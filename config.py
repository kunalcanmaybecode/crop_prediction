"""
Single Configuration File for Crop Classification System
Birla Institute of Technology, Mesra - B.Tech CSE Minor Project
"""

import os
from pathlib import Path

# Base Paths
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
LABELS_DIR = DATA_DIR / "labels"
RAW_DATA_DIR = DATA_DIR / "raw"
PROCESSED_DATA_DIR = DATA_DIR / "processed"
OUTPUTS_DIR = BASE_DIR / "outputs"
DOCS_DIR = BASE_DIR / "docs"
REPORT_DIR = BASE_DIR / "report"
APP_DIR = BASE_DIR / "app"

# Ensure runtime directories exist
for directory in [DATA_DIR, LABELS_DIR, RAW_DATA_DIR, PROCESSED_DATA_DIR, OUTPUTS_DIR, REPORT_DIR]:
    directory.mkdir(parents=True, exist_ok=True)

# Google Earth Engine Configuration
# User can set via environment variable GEE_PROJECT_ID or update below
GEE_PROJECT_ID = os.getenv("GEE_PROJECT_ID", "my-crop-classification")

# Study Region & Satellite Data
STUDY_STATE = "Uttar Pradesh"
# Sentinel-2 Harmonized Surface Reflectance
S2_COLLECTION = "COPERNICUS/S2_SR_HARMONIZED"
# Cloud Score+ Collection for Sentinel-2
CLOUD_SCORE_COLLECTION = "GOOGLE/CLOUD_SCORE_PLUS/V1/S2_HARMONIZED"
CLOUD_SCORE_BAND = "cs"  # or 'cs_cdf'
CLOUD_THRESHOLD = 0.60   # Minimum quality score (0 to 1) for clear pixels

# Temporal Window: AgriFieldNet Rabi Season (2021-2022)
# Covers typical vegetative, flowering, and ripening phenological stages
START_DATE = "2021-11-01"
END_DATE = "2022-04-30"

# Time Series Monthly Steps for Composites
TIME_STEPS = [
    ("2021-11-01", "2021-11-30", "2021_Nov"),
    ("2021-12-01", "2021-12-31", "2021_Dec"),
    ("2022-01-01", "2022-01-31", "2022_Jan"),
    ("2022-02-01", "2022-02-28", "2022_Feb"),
    ("2022-03-01", "2022-03-31", "2022_Mar"),
    ("2022-04-01", "2022-04-30", "2022_Apr"),
]

# Target Crop Classes (3 classes)
# AgriFieldNet official IDs: Wheat: 1, Mustard: 2, Maize: 9 (also supports normalized ID 3)
TARGET_CROPS = {
    1: "Wheat",
    2: "Mustard",
    9: "Maize",
    3: "Maize"  # Normalized class ID alias for Maize
}

# Crop name alias matching dictionary (for flexible CSV / Shapefile / GeoJSON columns)
CROP_NAME_MAPPING = {
    "wheat": "Wheat",
    "mustard": "Mustard",
    "maize": "Maize",
    "corn": "Maize",
    "sarson": "Mustard",
    "gehun": "Wheat",
    "gehu": "Wheat"
}

# Color palette for classes in visualizations and maps
CROP_COLORS = {
    "Wheat": "#2ca02c",    # Green
    "Mustard": "#e377c2",  # Yellowish Pink / Golden Yellow
    "Maize": "#ff7f0e"     # Orange
}

# Feature Engineering Settings
INDICES = ["NDVI", "EVI", "NDWI"]
# Sentinel-2 Band Names: B2=Blue, B3=Green, B4=Red, B8=NIR, B11=SWIR1
S2_BANDS = ["B2", "B3", "B4", "B8", "B11"]

# Inner Buffer for field polygons in meters (to remove edge/boundary mixed pixels)
FIELD_BUFFER_METERS = -5.0

# Model Training Parameters
RANDOM_STATE = 42
TEST_SIZE = 0.25
SPATIAL_SPLIT_COLUMN = "tile_id"  # Column for spatial splitting if available, else centroid spatial block

# Output Filenames
DATASET_STATS_CSV = OUTPUTS_DIR / "dataset_summary.csv"
METRICS_JSON = OUTPUTS_DIR / "metrics.json"
METRICS_CSV = OUTPUTS_DIR / "model_comparison.csv"
CONFUSION_MATRIX_PNG = OUTPUTS_DIR / "confusion_matrix.png"
FEATURE_IMPORTANCE_PNG = OUTPUTS_DIR / "feature_importance.png"
NDVI_PHENOLOGY_PNG = OUTPUTS_DIR / "crop_phenology_ndvi.png"
CROP_MAP_TIF = OUTPUTS_DIR / "crop_classification_map.tif"
CROP_MAP_PNG = OUTPUTS_DIR / "crop_classification_map.png"
OPTIONAL_POINTS_CSV = DATA_DIR / "my_points.csv"
