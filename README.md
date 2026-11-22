# Crop Classification System for India Using Sentinel-2 Time Series & Machine Learning

**Department of Computer Science and Engineering**  
**Birla Institute of Technology, Mesra, Ranchi, Jharkhand – 835215**  
**B.Tech Minor Project**

---

## 1. Project Overview
This project develops an end-to-end crop classification framework to identify **Wheat**, **Mustard**, and **Maize** in India (Uttar Pradesh) using multi-temporal **Sentinel-2 Harmonized Surface Reflectance** (`COPERNICUS/S2_SR_HARMONIZED`) satellite imagery, state-of-the-art cloud masking with **Cloud Score+** (`GOOGLE/CLOUD_SCORE_PLUS/V1/S2_HARMONIZED`), phenological vegetation indices (NDVI, EVI, NDWI), and machine learning classifiers (Random Forest, Support Vector Machine).

The system includes:
- Earth Engine data pipeline with Cloud Score+ cloud/shadow masking
- Phenological feature extraction per agricultural field
- Spatial cross-validation (field/tile-based, preventing spatial autocorrelation leakage)
- Interactive Streamlit dashboard for regional mapping and phenology analysis
- Standard department report conforming to BIT Mesra guidelines (`docs/Project  Report Template.pdf`)
- Presentation slides conforming to BIT Mesra presentation guidelines (`docs/Guidelines for Project Presentation.pdf`)

---

## 2. Directory Structure

```
crop_prediction/
├── config.py                 # Central configuration file for parameters and paths
├── requirements.txt          # Python dependencies
├── test_gee.py               # Earth Engine authentication & query verification test
├── README.md                 # Project documentation
├── app/                      # Streamlit interactive dashboard application
│   └── app.py
├── data/
│   ├── labels/               # Ground truth field polygons (AgriFieldNet India)
│   ├── raw/                  # Downloaded raw rasters / feature tables
│   ├── processed/            # Processed field time-series and train/test splits
│   └── my_points.csv         # (Optional) External validation point dataset
├── docs/                     # Department templates and guidelines
│   ├── Project  Report Template.pdf
│   └── Guidelines for Project Presentation.pdf
├── notebooks/                # Exploratory data analysis & experiments
├── outputs/                  # Final figures, metrics, confusion matrix, crop maps
│   ├── metrics.json
│   ├── model_comparison.csv
│   ├── dataset_summary.csv
│   ├── confusion_matrix.png
│   ├── feature_importance.png
│   ├── crop_phenology_ndvi.png
│   ├── crop_classification_map.png
│   └── crop_classification_map.tif
├── report/                   # Generated BIT Mesra DOCX report & PPTX slides
│   ├── Minor_Project_Report.docx
│   └── Minor_Project_Presentation.pptx
└── src/                      # Modular Python library
    ├── __init__.py
    ├── gee_utils.py          # Earth Engine extraction, cloud masking & compositing
    ├── features.py           # Phenological vegetation indices (NDVI, EVI, NDWI)
    ├── dataset.py            # Label loading, polygon buffering, spatial train/test split
    ├── models.py             # Random Forest & SVM training and evaluation
    └── mapping.py            # Regional classification map generation & export
```

---

## 3. Google Earth Engine Setup & Authentication

### Step 1: Create a Google Cloud Project with Earth Engine Enabled
If you do not already have a GEE Project ID:
1. Log in to your Google Account and visit:  
   **[https://code.earthengine.google.com/register](https://code.earthengine.google.com/register)**
2. Select **"Register a non-commercial or commercial Cloud Project"** (choose "Unpaid usage" -> "Academic & Research" for your university project).
3. Create a new Google Cloud Project (e.g., `ee-crop-classification` or your name/roll number) or select an existing project.
4. Note your **Google Cloud Project ID** (for example: `ee-crop-classification-2024` or `eminent-vector-123456`).
5. Ensure the **Earth Engine API** is enabled:  
   **[https://console.cloud.google.com/apis/library/earthengine.googleapis.com](https://console.cloud.google.com/apis/library/earthengine.googleapis.com)**

### Step 2: Authenticate Locally via Terminal
Run the Earth Engine authentication command in PowerShell:
```powershell
earthengine authenticate
```
A browser window will open asking you to sign in with your Google account. Grant the requested permissions, and Google will provide an authorization code or automatically store local credentials at `~/.config/earthengine/credentials`.

### Step 3: Configure Project ID & Run Verification Test
You can set your project ID as an environment variable in PowerShell:
```powershell
$env:GEE_PROJECT_ID = "YOUR_PROJECT_ID"
python test_gee.py
```
Or directly edit `GEE_PROJECT_ID` inside `config.py`.

When configured, `test_gee.py` will query Sentinel-2 Harmonized and Cloud Score+ over Uttar Pradesh and print:
```
>>> ALL GEE CHECKS PASSED SUCCESSFULLY! Ready for Stage 2. <<<
```

---

## 4. AgriFieldNet India Ground Truth Labels
Place your AgriFieldNet field polygon files into `data/labels/`.  
Supported formats:
- Shapefile archive (`.shp`, `.shx`, `.dbf`, `.prj`)
- GeoJSON (`.geojson` or `.json`)
- Geopackage (`.gpkg`)
- Or CSV with polygon WKT/geometry + crop label (`train_labels.csv` / `field_boundaries.geojson`)

Our pipeline automatically:
1. Filters records to **Wheat**, **Mustard**, and **Maize** (dropping all other crop classes).
2. Reprojects to UTM / metric CRS (`EPSG:32644` for Uttar Pradesh) and applies a **5 m inward buffer** to eliminate boundary mixed pixels.
3. Computes field-level median time-series for all composite periods.
4. Performs a spatial cross-validation train/test split.

---

## 5. Running the Pipeline
Once data and GEE credentials are in place:
```powershell
# Run data extraction and model training
python src/pipeline.py

# Launch the interactive Streamlit dashboard
streamlit run app/app.py
```
