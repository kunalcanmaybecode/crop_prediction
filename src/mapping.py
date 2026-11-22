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
    classified_array: np.ndarray,
    bounds: tuple,
    output_tif: Path = CROP_MAP_TIF,
    output_png: Path = CROP_MAP_PNG,
    crs: str = "EPSG:4326"
) -> None:
    """
    Saves the 2D classified crop array as a GeoTIFF and high-res PNG map.
    Class coding: 0: Background/Non-target, 1: Wheat, 2: Mustard, 3: Maize
    """
    height, width = classified_array.shape
    minx, miny, maxx, maxy = bounds
    transform = from_bounds(minx, miny, maxx, maxy, width, height)

    # 1. Export GeoTIFF
    print(f"[*] Writing GeoTIFF to {output_tif}...")
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
        dst.write(classified_array.astype(rasterio.uint8), 1)

    # 2. Export Publication-quality PNG Map
    print(f"[*] Generating cartographic PNG map to {output_png}...")
    cmap = ListedColormap(["#f0f0f0", CROP_COLORS["Wheat"], CROP_COLORS["Mustard"], CROP_COLORS["Maize"]])
    
    fig, ax = plt.subplots(figsize=(10, 8), dpi=300)
    plt.rcParams["font.family"] = "serif"
    plt.rcParams["font.serif"] = ["Times New Roman"] + plt.rcParams["font.serif"]

    im = ax.imshow(
        classified_array,
        cmap=cmap,
        extent=[minx, maxx, miny, maxy],
        origin="upper",
        vmin=0,
        vmax=3
    )

    ax.set_title("Sentinel-2 Crop Classification Map (Uttar Pradesh)", fontsize=14, weight="bold", pad=12)
    ax.set_xlabel("Longitude (°E)", fontsize=11, labelpad=8)
    ax.set_ylabel("Latitude (°N)", fontsize=11, labelpad=8)
    ax.ticklabel_format(useOffset=False, style='plain')
    ax.grid(True, linestyle=":", alpha=0.6)

    # Custom Legend
    legend_patches = [
        mpatches.Patch(color=CROP_COLORS["Wheat"], label="Wheat"),
        mpatches.Patch(color=CROP_COLORS["Mustard"], label="Mustard"),
        mpatches.Patch(color=CROP_COLORS["Maize"], label="Maize"),
        mpatches.Patch(color="#f0f0f0", edgecolor="grey", label="Other / Background")
    ]
    ax.legend(
        handles=legend_patches,
        loc="lower right",
        frameon=True,
        framealpha=0.9,
        facecolor="white",
        title="Crop Classes",
        title_fontsize=11,
        fontsize=10
    )

    plt.tight_layout()
    plt.savefig(output_png, dpi=300)
    plt.close()
    print(f"[+] Crop map successfully saved: GeoTIFF ({output_tif}) & PNG ({output_png})")


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
