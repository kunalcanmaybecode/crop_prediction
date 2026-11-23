"""
Google Earth Engine (GEE) Utilities for Sentinel-2 Processing
Birla Institute of Technology, Mesra - Minor Project

Implements:
1. Harmonized Sentinel-2 Surface Reflectance collection access
2. Cloud Score+ quality masking (GOOGLE/CLOUD_SCORE_PLUS/V1/S2_HARMONIZED)
3. Multi-temporal monthly median compositing
"""

import sys
from pathlib import Path

# Add project root to sys.path
PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import ee
from config import (
    GEE_PROJECT_ID,
    S2_COLLECTION,
    CLOUD_SCORE_COLLECTION,
    CLOUD_SCORE_BAND,
    CLOUD_THRESHOLD,
    TIME_STEPS,
)


def initialize_gee(project_id: str = None) -> bool:
    """Initialize Earth Engine with the designated Google Cloud Project."""
    pid = project_id or GEE_PROJECT_ID
    if not pid or pid == "<MY_GEE_PROJECT_ID>":
        raise ValueError(
            "GEE_PROJECT_ID is not configured. Please set GEE_PROJECT_ID in config.py "
            "or as an environment variable."
        )
    ee.Initialize(project=pid)
    return True


def mask_s2_clouds(image: ee.Image) -> ee.Image:
    """
    Cloud and cloud-shadow masking using Cloud Score+.
    
    Why Cloud Masking Matters in Remote Sensing & Crop Phenology:
    1. Clouds falsely inflate visible reflectance and skew vegetation indices.
    2. Cloud shadows drastically depress NIR and SWIR signals, causing artificial
       drops in NDVI and EVI that mimic plant disease, flooding, or harvest.
    3. Cloud Score+ assesses both clear-sky probability ('cs') and cumulative 
       distribution ('cs_cdf'), giving cleaner separation than legacy QA60 bitmask.
    """
    # Cloud Score+ QA image linked to this Sentinel-2 scene
    # Range is 0 (cloud/shadow) to 1 (clear sky)
    cs = image.select(CLOUD_SCORE_BAND)
    clear_mask = cs.gte(CLOUD_THRESHOLD)
    
    # Scale optical bands to reflectance (0.0 - 1.0)
    # Sentinel-2 L2A Harmonized stores reflectance scaled by 10000
    optical_bands = ["B2", "B3", "B4", "B8", "B11"]
    scaled = image.select(optical_bands).divide(10000.0)
    
    return scaled.updateMask(clear_mask).copyProperties(image, ["system:time_start"])


def get_s2_composite(
    roi: ee.Geometry,
    start_date: str,
    end_date: str
) -> ee.Image:
    """
    Fetch and composite Sentinel-2 Harmonized scenes for a time window and region.
    Uses Google Cloud Score+ QA layer to produce high-quality, cloud-free composite.
    Uses qualityMosaic('cs') which picks the highest clear-sky probability pixel
    for every location, preventing persistent winter fog / haze data gaps in North India.
    """
    s2 = (
        ee.ImageCollection(S2_COLLECTION)
        .filterBounds(roi)
        .filterDate(start_date, end_date)
    )
    
    cs_plus = (
        ee.ImageCollection(CLOUD_SCORE_COLLECTION)
        .filterBounds(roi)
        .filterDate(start_date, end_date)
    )
    
    optical_bands = ["B2", "B3", "B4", "B8", "B11"]
    
    # Link Sentinel-2 scenes with their Cloud Score+ QA band
    s2_with_cs = s2.linkCollection(cs_plus, [CLOUD_SCORE_BAND])
    
    # Scale optical bands to reflectance (0.0 to 1.0) and retain QA score
    def scale_and_retain_qa(img):
        scaled = img.select(optical_bands).divide(10000.0)
        return scaled.addBands(img.select(CLOUD_SCORE_BAND))
        
    scaled_coll = s2_with_cs.map(scale_and_retain_qa)
    
    # Quality mosaic selects the observation with the maximum clear sky score ('cs')
    # Note: Avoid eager .clip(roi) across massive states to prevent unnecessary geometry intersection overhead
    composite = scaled_coll.qualityMosaic(CLOUD_SCORE_BAND).select(optical_bands)
    return composite


def build_seasonal_composites(roi: ee.Geometry) -> dict:
    """
    Constructs multi-temporal monthly median composites across the Rabi growing season.
    Returns a dictionary of {period_name: ee.Image}.
    """
    composites = {}
    for start, end, label in TIME_STEPS:
        comp = get_s2_composite(roi, start, end)
        composites[label] = comp
    return composites
