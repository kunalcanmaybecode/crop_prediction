"""
Crop Classification Map Generation and GeoTIFF Export
Birla Institute of Technology, Mesra - Minor Project

Produces:
1. Color-coded classification raster GeoTIFF
2. Publication-quality PNG map with legend and scale
3. geemap interactive visualization helper
"""

import sys
from pathlib import Path

# Add project root to sys.path
PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import numpy as np
import geopandas as gpd
import matplotlib.pyplot as plt
from matplotlib.colors import ListedColormap
import matplotlib.patches as mpatches
import rasterio
from rasterio.transform import from_bounds
import geemap
import ee

from config import (
    CROP_MAP_TIF,
    CROP_MAP_PNG,
    CROP_COLORS,
    OUTPUTS_DIR
)


def export_classification_map(
    gdf_classified: gpd.GeoDataFrame,
    bounds: tuple = None,
    output_tif: Path = CROP_MAP_TIF,
    output_png: Path = CROP_MAP_PNG,
    crs: str = "EPSG:4326"
) -> None:
    """
    Saves a genuine, georeferenced spatial classification raster GeoTIFF
    and publication-quality cartographic PNG map over the core Uttar Pradesh study area.
    Class coding: 0: Background/Non-target, 1: Wheat, 2: Mustard, 3: Maize
    """
    from rasterio.features import rasterize
    from shapely.geometry import box
    
    # Focus on the core agricultural cluster in UP (Bahraich-Gonda corridor)
    # where Wheat, Mustard, and Maize are all densely present
    if bounds is None:
        minx, miny, maxx, maxy = 82.15, 27.35, 27.75, 27.65
    else:
        minx, miny, maxx, maxy = bounds

    # Grid dimensions (approx. 20m spatial resolution across the zone)
    width = 1200
    height = 900
    transform = from_bounds(minx, miny, maxx, maxy, width, height)

    # Filter fields within or intersecting this study area
    study_box = box(minx, miny, maxx, maxy)
    roi_fields = gdf_classified[gdf_classified.geometry.intersects(study_box)].copy()
    
    crop_int_map = {"Wheat": 1, "Mustard": 2, "Maize": 3}
    
    shapes_to_rasterize = []
    for _, row in roi_fields.iterrows():
        c_code = crop_int_map.get(row.get("crop_name", ""), 0)
        if c_code > 0 and row.geometry.is_valid:
            shapes_to_rasterize.append((row.geometry, c_code))

    # Rasterize field polygons with their predicted classes
    if shapes_to_rasterize:
        raster_data = rasterize(
            shapes=shapes_to_rasterize,
            out_shape=(height, width),
            transform=transform,
            fill=0,
            dtype=np.uint8
        )
    else:
        raster_data = np.zeros((height, width), dtype=np.uint8)

    # 1. Export GeoTIFF
    print(f"[*] Writing georeferenced GeoTIFF to {output_tif}...")
    with rasterio.open(
        output_tif,
        "w",
        driver="GTiff",
        height=height,
        width=width,
        count=1,
        dtype=rasterio.uint8,
        crs=crs,
        transform=transform,
        nodata=0
    ) as dst:
        dst.write(raster_data, 1)

    # 2. Export Publication-quality Cartographic PNG Map
    print(f"[*] Generating cartographic PNG map to {output_png}...")
    cmap = ListedColormap(["#f5f7fa", CROP_COLORS["Wheat"], CROP_COLORS["Mustard"], CROP_COLORS["Maize"]])
    
    fig, ax = plt.subplots(figsize=(12, 8), dpi=300)
    plt.rcParams["font.family"] = "serif"
    plt.rcParams["font.serif"] = ["Times New Roman"] + plt.rcParams["font.serif"]

    im = ax.imshow(
        raster_data,
        cmap=cmap,
        extent=[minx, maxx, miny, maxy],
        origin="upper",
        vmin=0,
        vmax=3,
        interpolation="nearest"
    )

    # Plot field vector boundaries for high-resolution clarity
    for _, row in roi_fields.iterrows():
        c = row.get("crop_name", "")
        color = CROP_COLORS.get(c, "#999999")
        if row.geometry.geom_type == "Polygon":
            x, y = row.geometry.exterior.xy
            ax.plot(x, y, color=color, linewidth=1.2, alpha=0.9)
            ax.fill(x, y, color=color, alpha=0.7)
        elif row.geometry.geom_type == "MultiPolygon":
            for poly in row.geometry.geoms:
                x, y = poly.exterior.xy
                ax.plot(x, y, color=color, linewidth=1.2, alpha=0.9)
                ax.fill(x, y, color=color, alpha=0.7)

    # Calculate field statistics for this region
    n_wheat = (roi_fields["crop_name"] == "Wheat").sum()
    n_mustard = (roi_fields["crop_name"] == "Mustard").sum()
    n_maize = (roi_fields["crop_name"] == "Maize").sum()

    ax.set_title(
        "Sentinel-2 Crop Classification Map — Uttar Pradesh, India\n"
        "(Bahraich – Gonda Agricultural Corridor: 27.35°N – 27.65°N, 82.15°E – 82.75°E)",
        fontsize=13,
        weight="bold",
        pad=14
    )
    ax.set_xlabel("Longitude (°E)", fontsize=11, labelpad=8)
    ax.set_ylabel("Latitude (°N)", fontsize=11, labelpad=8)
    ax.ticklabel_format(useOffset=False, style='plain')
    ax.grid(True, linestyle=":", alpha=0.6, color="#888888")

    # Add North Arrow
    ax.text(
        0.96, 0.93, "▲\nN",
        transform=ax.transAxes,
        fontsize=13,
        weight="bold",
        ha="center",
        va="center",
        bbox=dict(boxstyle="square,pad=0.3", facecolor="white", edgecolor="black", alpha=0.85)
    )

    # Add Scale bar text
    ax.text(
        0.04, 0.05, "Scale: 1 : 50,000 | Datum: WGS84 | Sensor: Sentinel-2 Harmonized (10 m)",
        transform=ax.transAxes,
        fontsize=9,
        style="italic",
        bbox=dict(boxstyle="round,pad=0.3", facecolor="white", edgecolor="grey", alpha=0.85)
    )

    # Custom Detailed Legend
    legend_patches = [
        mpatches.Patch(color=CROP_COLORS["Wheat"], label=f"Wheat (n = {n_wheat} fields)"),
        mpatches.Patch(color=CROP_COLORS["Mustard"], label=f"Mustard (n = {n_mustard} fields)"),
        mpatches.Patch(color=CROP_COLORS["Maize"], label=f"Maize (n = {n_maize} fields)"),
        mpatches.Patch(color="#f5f7fa", edgecolor="grey", label="Background / Non-agricultural")
    ]
    ax.legend(
        handles=legend_patches,
        loc="lower right",
        frameon=True,
        framealpha=0.92,
        facecolor="white",
        edgecolor="#cccccc",
        title="Classified Crop Classes",
        title_fontsize=11,
        fontsize=10
    )

    plt.tight_layout()
    plt.savefig(output_png, dpi=300)
    plt.close()
    print(f"[+] Spatial crop map successfully saved: GeoTIFF ({output_tif}) & PNG ({output_png})")


def create_geemap_view(roi_bounds: tuple, map_center: list = [26.85, 80.95], zoom: int = 10):
    """Creates a geemap Map instance with Sentinel-2 and classification layers."""
    Map = geemap.Map(center=map_center, zoom=zoom)
    # Add Google Satellite basemap
    Map.add_basemap("SATELLITE")
    
    # Legend dictionary
    legend_dict = {
        "Wheat": CROP_COLORS["Wheat"],
        "Mustard": CROP_COLORS["Mustard"],
        "Maize": CROP_COLORS["Maize"],
    }
    Map.add_legend(title="Crop Classification", legend_dict=legend_dict)
    return Map
