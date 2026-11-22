"""
Automated Downloader & Vectorizer for AgriFieldNet India Labels
Birla Institute of Technology, Mesra - Minor Project

Downloads label chips directly from Source Cooperative HTTP endpoint,
extracts ground truth agricultural field boundaries as vector polygons,
filters to target classes (Wheat: 1, Mustard: 2, Maize: 9),
and saves directly to data/labels/agrifieldnet_fields.geojson.
"""

import os
import io
import time
import urllib.request
import xml.etree.ElementTree as ET
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed

import numpy as np
import pandas as pd
import geopandas as gpd
import rasterio
from rasterio.features import shapes
from shapely.geometry import shape

from config import LABELS_DIR, TARGET_CROPS

BASE_URL = "https://data.source.coop/radiantearth/agrifieldnet-competition/"


def get_all_train_label_keys(max_tiles: int = 150):
    """
    Fetch S3 keys for train_labels from Source Cooperative.
    Pairs each tile's field_ids.tif and label .tif.
    """
    print("[*] Querying Source Cooperative catalog for AgriFieldNet training labels...")
    url = f"{BASE_URL}?prefix=train_labels/"
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    
    with urllib.request.urlopen(req) as resp:
        root = ET.fromstring(resp.read())
        ns = {"s3": "http://s3.amazonaws.com/doc/2006-03-01/"}
        keys = [k.text for k in root.findall(".//s3:Key", ns)]

    # Group into pairs: {chip_id: {'label': key, 'field_ids': key}}
    tiles = {}
    for k in keys:
        fname = k.split("/")[-1]
        if not fname.endswith(".tif"):
            continue
        if "_field_ids.tif" in fname:
            chip_id = fname.replace("_field_ids.tif", "").replace("ref_agrifieldnet_competition_v1_labels_train_", "")
            tiles.setdefault(chip_id, {})["field_ids"] = k
        else:
            chip_id = fname.replace(".tif", "").replace("ref_agrifieldnet_competition_v1_labels_train_", "")
            tiles.setdefault(chip_id, {})["label"] = k

    # Filter to tiles that have both files
    complete_tiles = {cid: paths for cid, paths in tiles.items() if "label" in paths and "field_ids" in paths}
    print(f"[+] Found {len(complete_tiles)} complete training tiles in catalog.")
    
    # Select subset or all (up to max_tiles to keep runtime fast and focused)
    selected_cids = list(complete_tiles.keys())[:max_tiles]
    return {cid: complete_tiles[cid] for cid in selected_cids}


def download_and_vectorize_tile(chip_id: str, paths: dict):
    """
    Fetch both rasters in memory, extract field polygons, and return records.
    """
    # Key in S3: agrifieldnet-competition/train_labels/...
    # URL path: https://data.source.coop/radiantearth/agrifieldnet-competition/train_labels/...
    label_path = paths["label"].replace("agrifieldnet-competition/", "")
    fields_path = paths["field_ids"].replace("agrifieldnet-competition/", "")
    
    url_label = f"{BASE_URL}{label_path}"
    url_fields = f"{BASE_URL}{fields_path}"
    
    try:
        req_l = urllib.request.Request(url_label, headers={"User-Agent": "Mozilla/5.0"})
        req_f = urllib.request.Request(url_fields, headers={"User-Agent": "Mozilla/5.0"})
        
        with urllib.request.urlopen(req_l, timeout=15) as rl, urllib.request.urlopen(req_f, timeout=15) as rf:
            b_label = io.BytesIO(rl.read())
            b_fields = io.BytesIO(rf.read())
            
        with rasterio.open(b_label) as ds_label, rasterio.open(b_fields) as ds_fields:
            arr_label = ds_label.read(1)
            arr_fields = ds_fields.read(1)
            transform = ds_label.transform
            crs = ds_label.crs
            
            records = []
            unique_fields = np.unique(arr_fields)
            
            for fid in unique_fields:
                if fid == 0:
                    continue
                mask = (arr_fields == fid)
                field_crops = arr_label[mask]
                if len(field_crops) == 0:
                    continue
                crop_id = int(pd.Series(field_crops).mode()[0])
                
                # Check if target crop (1: Wheat, 2: Mustard, 9: Maize, 3: Maize alias)
                if crop_id not in TARGET_CROPS:
                    continue
                    
                crop_name = TARGET_CROPS[crop_id]
                
                field_mask = mask.astype(np.uint8)
                for geom, val in shapes(field_mask, mask=field_mask == 1, transform=transform):
                    poly = shape(geom)
                    if poly.is_valid and not poly.is_empty:
                        records.append({
                            "field_id": int(fid),
                            "crop_id": crop_id,
                            "crop_name": crop_name,
                            "tile_id": chip_id,
                            "geometry": poly,
                            "crs": str(crs)
                        })
            return records
    except Exception as e:
        return []


def download_agrifieldnet_labels(output_geojson: Path = LABELS_DIR / "agrifieldnet_fields.geojson", max_tiles: int = 150):
    """
    Main function to parallel download label chips and generate vector GeoJSON.
    """
    tile_dict = get_all_train_label_keys(max_tiles=max_tiles)
    print(f"[*] Downloading and processing {len(tile_dict)} label tiles in parallel...")
    
    all_records = []
    start_time = time.time()
    
    with ThreadPoolExecutor(max_workers=10) as executor:
        futures = {executor.submit(download_and_vectorize_tile, cid, paths): cid for cid, paths in tile_dict.items()}
        for future in as_completed(futures):
            recs = future.result()
            if recs:
                all_records.extend(recs)
                
    elapsed = time.time() - start_time
    print(f"[+] Download and vectorization completed in {elapsed:.1f}s.")
    print(f"[+] Total raw field polygons extracted: {len(all_records)}")
    
    if not all_records:
        raise RuntimeError("No fields could be extracted from the label tiles.")

    # Group by CRS and convert to WGS84 (EPSG:4326)
    gdfs = []
    df = pd.DataFrame(all_records)
    for crs_val, group in df.groupby("crs"):
        gdf_group = gpd.GeoDataFrame(group, geometry="geometry", crs=crs_val)
        gdf_wgs84 = gdf_group.to_crs(epsg=4326)
        gdfs.append(gdf_wgs84)
        
    combined_gdf = gpd.GeoDataFrame(pd.concat(gdfs, ignore_index=True), crs="EPSG:4326")
    
    # Deduplicate field_ids if any field crossed tile boundaries
    combined_gdf = combined_gdf.dissolve(by="field_id", aggfunc="first").reset_index()

    # Filter to bounding box of Uttar Pradesh / Northern India agricultural belt
    # Uttar Pradesh coordinates approximately: Lon 77.0 to 84.5 E, Lat 23.8 to 30.5 N
    up_gdf = combined_gdf[
        (combined_gdf.geometry.centroid.x >= 77.0) & (combined_gdf.geometry.centroid.x <= 84.5) &
        (combined_gdf.geometry.centroid.y >= 24.0) & (combined_gdf.geometry.centroid.y <= 30.5)
    ].copy()
    
    target_gdf = up_gdf if len(up_gdf) >= 30 else combined_gdf
    
    # Save GeoJSON
    target_gdf.to_file(output_geojson, driver="GeoJSON")
    print(f"[+] Successfully saved {len(target_gdf)} ground truth field polygons to:\n    {output_geojson}")
    print("\nClass breakdown:")
    print(target_gdf["crop_name"].value_counts().to_string())
    return target_gdf


if __name__ == "__main__":
    download_agrifieldnet_labels()
