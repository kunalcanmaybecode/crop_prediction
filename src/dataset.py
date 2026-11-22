"""
AgriFieldNet Ground Truth Processing & Dataset Preparation
Birla Institute of Technology, Mesra - Minor Project

Handles:
1. Loading field polygons (GeoJSON/Shapefile/CSV)
2. Filtering to target classes: Wheat, Mustard, Maize
3. CRS metric transformation (EPSG:32644) and 5 m inward buffering
4. Spatial train/test splitting (by tile or spatial block centroid)
5. Dataset summary statistics export
"""

import os
import sys
from pathlib import Path

# Add project root to sys.path
PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import numpy as np
import pandas as pd
import geopandas as gpd
from shapely.geometry import shape
from shapely import wkt
from sklearn.model_selection import GroupKFold, train_test_split

from config import (
    LABELS_DIR,
    PROCESSED_DATA_DIR,
    TARGET_CROPS,
    CROP_NAME_MAPPING,
    FIELD_BUFFER_METERS,
    RANDOM_STATE,
    TEST_SIZE,
    DATASET_STATS_CSV,
    OPTIONAL_POINTS_CSV
)


def find_label_files(labels_dir: Path = LABELS_DIR):
    """Detect ground truth vector or raster chip label files in the labels directory."""
    patterns = ["*.geojson", "*.json", "*.shp", "*.gpkg", "*.csv"]
    files = []
    for p in patterns:
        files.extend(list(labels_dir.glob(p)))
    # Also check for raster chip folders containing field_ids.tif or raster_labels.tif
    train_label_folders = list(labels_dir.glob("**/raster_labels.tif"))
    if train_label_folders:
        files.append(labels_dir)
    return files


def load_from_raster_chips(labels_dir: Path) -> gpd.GeoDataFrame:
    """
    Extract vector field polygons from AgriFieldNet raster chip pairs
    (field_ids.tif and raster_labels.tif) as downloaded from Source Cooperative.
    """
    import rasterio
    from rasterio.features import shapes
    
    print("[*] Parsing AgriFieldNet raster label chips from Source Cooperative...")
    label_files = list(labels_dir.glob("**/raster_labels.tif"))
    if not label_files:
        raise ValueError(f"No raster_labels.tif chips found in {labels_dir}")
        
    records = []
    for label_path in label_files:
        chip_folder = label_path.parent
        field_id_path = chip_folder / "field_ids.tif"
        if not field_id_path.exists():
            continue
            
        with rasterio.open(label_path) as src_label, rasterio.open(field_id_path) as src_field:
            labels_arr = src_label.read(1)
            field_arr = src_field.read(1)
            transform = src_label.transform
            crs = src_label.crs
            
            # Find unique fields in this tile
            unique_fields = np.unique(field_arr)
            for fid in unique_fields:
                if fid == 0:
                    continue
                mask = (field_arr == fid)
                # Crop label for this field
                field_crop_vals = labels_arr[mask]
                if len(field_crop_vals) == 0:
                    continue
                crop_id = int(pd.Series(field_crop_vals).mode()[0])
                
                # Check if target crop (1=Wheat, 2=Mustard, 9=Maize, 3=Maize alias)
                if crop_id not in TARGET_CROPS:
                    continue
                    
                crop_name = TARGET_CROPS[crop_id]
                
                # Vectorize field boundary
                field_mask = mask.astype(np.uint8)
                for geom, val in shapes(field_mask, mask=field_mask == 1, transform=transform):
                    poly = shape(geom)
                    records.append({
                        "field_id": int(fid),
                        "crop_id": crop_id,
                        "crop_name": crop_name,
                        "tile_id": chip_folder.name,
                        "geometry": poly,
                        "crs": str(crs)
                    })

    if not records:
        raise ValueError("No matching fields found for target crops in raster chips.")
        
    gdf = gpd.GeoDataFrame(records, geometry="geometry", crs=records[0]["crs"])
    gdf = gdf.to_crs(epsg=4326)
    print(f"[+] Vectorized {len(gdf)} field polygons from raster chips.")
    return gdf


def load_and_filter_labels(file_path: Path) -> gpd.GeoDataFrame:
    """
    Load vector layer or raster chips with geopandas, harmonize crop names to target 3 classes,
    and drop unwanted crop classes.
    """
    if file_path.is_dir():
        return load_from_raster_chips(file_path)

    print(f"[*] Reading ground truth labels from {file_path}...")
    if file_path.suffix.lower() == ".csv":
        df = pd.read_csv(file_path)
        # Look for geometry or WKT column
        geom_col = next((c for c in df.columns if c.lower() in ["geometry", "wkt", "geom"]), None)
        if geom_col:
            df["geometry"] = df[geom_col].apply(wkt.loads)
            gdf = gpd.GeoDataFrame(df, geometry="geometry", crs="EPSG:4326")
        elif "latitude" in df.columns and "longitude" in df.columns:
            gdf = gpd.GeoDataFrame(
                df,
                geometry=gpd.points_from_xy(df.longitude, df.latitude),
                crs="EPSG:4326"
            )
        else:
            raise ValueError(f"CSV file {file_path} has no geometry or latitude/longitude columns.")
    else:
        gdf = gpd.read_file(file_path)

    # Standardize column names
    crop_col = next((c for c in gdf.columns if c.lower() in ["crop", "crop_name", "crop_type", "label", "crop_id"]), None)
    if not crop_col:
        raise ValueError(f"Could not identify crop label column in {file_path}. Columns: {list(gdf.columns)}")

    # Harmonize crop names to Wheat, Mustard, Maize
    def standardize_crop(val):
        if pd.isna(val):
            return None
        # If integer crop id
        if isinstance(val, (int, np.integer)) or (isinstance(val, float) and val.is_integer()):
            val_int = int(val)
            if val_int in TARGET_CROPS:
                return TARGET_CROPS[val_int]
        # If string
        s = str(val).strip().lower()
        if s in CROP_NAME_MAPPING:
            return CROP_NAME_MAPPING[s]
        for key in CROP_NAME_MAPPING:
            if key in s:
                return CROP_NAME_MAPPING[key]
        return None

    gdf["standard_crop"] = gdf[crop_col].apply(standardize_crop)
    
    # Filter to only the 3 target crops (Wheat, Mustard, Maize)
    filtered = gdf[gdf["standard_crop"].notna()].copy()
    filtered["crop_name"] = filtered["standard_crop"]
    print(f"[+] Loaded {len(filtered)} valid fields for target crops:")
    print(filtered["crop_name"].value_counts().to_string())
    
    return filtered


def buffer_field_polygons(gdf: gpd.GeoDataFrame, buffer_meters: float = FIELD_BUFFER_METERS) -> gpd.GeoDataFrame:
    """
    Project to UTM (EPSG:32644 for Uttar Pradesh) and apply negative buffer
    to eliminate mixed boundary pixels (edges where trees, paths, or borders mix with crop).
    Fallback to original geometry if inner buffer collapses small fields.
    """
    print(f"[*] Reprojecting to UTM (EPSG:32644) and applying {buffer_meters} m inner buffer...")
    gdf_utm = gdf.to_crs(epsg=32644)
    
    # Calculate original area in hectares
    gdf_utm["area_ha"] = gdf_utm.geometry.area / 10000.0
    
    # Apply negative buffer
    buffered_geom = gdf_utm.geometry.buffer(buffer_meters)
    
    # Check for empty geometries (where field was smaller than 10 m across)
    too_small = buffered_geom.is_empty
    if too_small.any():
        print(f"[!] Note: {too_small.sum()} small fields collapsed with {buffer_meters} m buffer; retaining unbuffered boundary.")
        buffered_geom = buffered_geom.where(~too_small, gdf_utm.geometry)
        
    gdf_utm["geometry"] = buffered_geom
    
    # Project back to WGS84 for Earth Engine compatibility
    buffered_wgs84 = gdf_utm.to_crs(epsg=4326)
    buffered_wgs84["area_ha"] = gdf_utm["area_ha"]
    return buffered_wgs84


def spatial_train_test_split(gdf: gpd.GeoDataFrame, test_size: float = TEST_SIZE, random_state: int = RANDOM_STATE):
    """
    Perform a strict spatial train/test split.
    Prevents spatial autocorrelation leakage: fields from the same spatial group/tile
    are never split across train and test sets.
    """
    # Check if tile_id or spatial group column exists
    tile_col = next((c for c in gdf.columns if c.lower() in ["tile_id", "tile", "grid_id", "cluster_id"]), None)
    
    if tile_col and gdf[tile_col].nunique() > 3:
        print(f"[*] Splitting by spatial tile column: '{tile_col}'...")
        tiles = gdf[tile_col].unique()
        train_tiles, test_tiles = train_test_split(tiles, test_size=test_size, random_state=random_state)
        train_gdf = gdf[gdf[tile_col].isin(train_tiles)].copy()
        test_gdf = gdf[gdf[tile_col].isin(test_tiles)].copy()
    else:
        print("[*] Creating spatial grid blocks from field centroids for spatial split...")
        centroids = gdf.geometry.centroid
        lons = centroids.x
        lats = centroids.y
        # 5x5 spatial grid blocks
        lon_bins = pd.cut(lons, bins=5, labels=False)
        lat_bins = pd.cut(lats, bins=5, labels=False)
        gdf["spatial_block"] = lon_bins.astype(str) + "_" + lat_bins.astype(str)
        
        blocks = gdf["spatial_block"].unique()
        train_blocks, test_blocks = train_test_split(blocks, test_size=test_size, random_state=random_state)
        train_gdf = gdf[gdf["spatial_block"].isin(train_blocks)].copy()
        test_gdf = gdf[gdf["spatial_block"].isin(test_blocks)].copy()
        
    print(f"[+] Spatial Train size: {len(train_gdf)} fields | Test size: {len(test_gdf)} fields")
    return train_gdf, test_gdf


def generate_dataset_summary(gdf: gpd.GeoDataFrame, train_gdf: gpd.GeoDataFrame, test_gdf: gpd.GeoDataFrame) -> pd.DataFrame:
    """
    Save annotated dataset description table (counts, area statistics, split sizes)
    to outputs/dataset_summary.csv for project report Table in Chapter 3.
    """
    summary_rows = []
    crops = ["Wheat", "Mustard", "Maize"]
    
    for crop in crops:
        total_c = gdf[gdf["crop_name"] == crop]
        train_c = train_gdf[train_gdf["crop_name"] == crop]
        test_c = test_gdf[test_gdf["crop_name"] == crop]
        
        n_total = len(total_c)
        n_train = len(train_c)
        n_test = len(test_c)
        
        area_mean = total_c["area_ha"].mean() if "area_ha" in total_c.columns and n_total > 0 else 0.0
        area_min = total_c["area_ha"].min() if "area_ha" in total_c.columns and n_total > 0 else 0.0
        area_max = total_c["area_ha"].max() if "area_ha" in total_c.columns and n_total > 0 else 0.0
        area_total = total_c["area_ha"].sum() if "area_ha" in total_c.columns and n_total > 0 else 0.0
        
        summary_rows.append({
            "Crop Class": crop,
            "Total Fields": n_total,
            "Train Fields": n_train,
            "Test Fields": n_test,
            "Mean Area (ha)": round(area_mean, 3),
            "Min Area (ha)": round(area_min, 3),
            "Max Area (ha)": round(area_max, 3),
            "Total Area (ha)": round(area_total, 2)
        })
        
    summary_df = pd.DataFrame(summary_rows)
    summary_df.to_csv(DATASET_STATS_CSV, index=False)
    print(f"[+] Dataset summary saved to {DATASET_STATS_CSV}")
    return summary_df
