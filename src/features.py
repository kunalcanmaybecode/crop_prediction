"""
Feature Engineering: Phenological Vegetation Indices
Birla Institute of Technology, Mesra - Minor Project

Computes:
1. NDVI (Normalized Difference Vegetation Index) - Canopy greenness & biomass
2. EVI (Enhanced Vegetation Index) - Biomass in dense canopies (resists saturation)
3. NDWI (Normalized Difference Water Index) - Canopy moisture / water content
And generates phenological curves for crop separation.
"""

import sys
from pathlib import Path

# Add project root to sys.path
PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import ee
from config import CROP_COLORS, NDVI_PHENOLOGY_PNG, TIME_STEPS


def compute_spectral_indices_ee(image: ee.Image) -> ee.Image:
    """
    Compute NDVI, EVI, and NDWI on an Earth Engine Sentinel-2 reflectance image.
    Bands expected (0.0 to 1.0): B2 (Blue), B3 (Green), B4 (Red), B8 (NIR), B11 (SWIR1).
    """
    # 1. NDVI = (NIR - Red) / (NIR + Red)
    ndvi = image.normalizedDifference(["B8", "B4"]).rename("NDVI")
    
    # 2. EVI = 2.5 * (NIR - Red) / (NIR + 6 * Red - 7.5 * Blue + 1)
    evi = image.expression(
        "2.5 * ((NIR - RED) / (NIR + 6.0 * RED - 7.5 * BLUE + 1.0))",
        {
            "NIR": image.select("B8"),
            "RED": image.select("B4"),
            "BLUE": image.select("B2"),
        }
    ).rename("EVI")
    
    # 3. NDWI (Gao 1996 / LSWI) = (NIR - SWIR1) / (NIR + SWIR1)
    ndwi = image.normalizedDifference(["B8", "B11"]).rename("NDWI")
    
    return image.addBands([ndvi, evi, ndwi])


def compute_spectral_indices_np(df: pd.DataFrame, time_labels: list) -> pd.DataFrame:
    """
    Compute indices from tabular reflectance values if raw band reflectances are extracted.
    """
    out = df.copy()
    for t in time_labels:
        b8 = out[f"{t}_B8"]
        b4 = out[f"{t}_B4"]
        b2 = out[f"{t}_B2"]
        b11 = out[f"{t}_B11"]
        
        # NDVI
        denom_ndvi = b8 + b4
        denom_ndvi = np.where(denom_ndvi == 0, 1e-6, denom_ndvi)
        out[f"{t}_NDVI"] = (b8 - b4) / denom_ndvi
        
        # EVI
        denom_evi = b8 + 6.0 * b4 - 7.5 * b2 + 1.0
        denom_evi = np.where(denom_evi == 0, 1e-6, denom_evi)
        out[f"{t}_EVI"] = 2.5 * (b8 - b4) / denom_evi
        
        # NDWI
        denom_ndwi = b8 + b11
        denom_ndwi = np.where(denom_ndwi == 0, 1e-6, denom_ndwi)
        out[f"{t}_NDWI"] = (b8 - b11) / denom_ndwi
    return out


def plot_ndvi_phenology(
    df: pd.DataFrame,
    crop_col: str = "crop_name",
    output_path: str = NDVI_PHENOLOGY_PNG
) -> None:
    """
    Plot mean NDVI time series curves per crop class with standard error shading.
    Explains the distinctive phenological signatures of Wheat, Mustard, and Maize.
    """
    ndvi_cols = [f"{step[2]}_NDVI" for step in TIME_STEPS]
    time_names = [step[2].replace("_", " ") for step in TIME_STEPS]
    
    plt.figure(figsize=(10, 6))
    plt.rcParams["font.family"] = "serif"
    plt.rcParams["font.serif"] = ["Times New Roman"] + plt.rcParams["font.serif"]

    for crop, color in CROP_COLORS.items():
        crop_subset = df[df[crop_col].str.lower() == crop.lower()]
        if crop_subset.empty:
            continue
        
        # Filter valid NDVI columns present in data
        valid_cols = [c for c in ndvi_cols if c in crop_subset.columns]
        if not valid_cols:
            continue
            
        mean_vals = crop_subset[valid_cols].mean().values
        std_vals = crop_subset[valid_cols].std().values
        n = len(crop_subset)
        sem_vals = std_vals / np.sqrt(max(n, 1))

        x_coords = np.arange(len(valid_cols))
        plt.plot(x_coords, mean_vals, marker="o", linewidth=2.5, label=f"{crop} (n={n})", color=color)
        plt.fill_between(x_coords, mean_vals - sem_vals, mean_vals + sem_vals, color=color, alpha=0.18)

    plt.title("Sentinel-2 NDVI Phenological Trajectories (Rabi Season 2021-2022)", fontsize=14, weight="bold", pad=12)
    plt.xlabel("Observation Period", fontsize=12, labelpad=8)
    plt.ylabel("Normalized Difference Vegetation Index (NDVI)", fontsize=12, labelpad=8)
    plt.xticks(np.arange(len(time_names)), time_names, fontsize=11)
    plt.yticks(fontsize=11)
    plt.ylim(0.0, 1.0)
    plt.grid(True, linestyle="--", alpha=0.5)
    plt.legend(frameon=True, fontsize=11, loc="upper left")
    plt.tight_layout()
    
    plt.savefig(output_path, dpi=300)
    plt.close()
    print(f"[+] Phenology chart saved to: {output_path}")
