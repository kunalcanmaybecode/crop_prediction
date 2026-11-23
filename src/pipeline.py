"""
End-to-End Execution Pipeline
Birla Institute of Technology, Mesra - Minor Project

Orchestrates:
1. GEE initialization
2. AgriFieldNet label ingestion & spatial buffering
3. Sentinel-2 multi-temporal extraction & cloud masking
4. Feature engineering (NDVI, EVI, NDWI) & phenology plotting
5. Spatial train/test split & dataset summary export
6. Random Forest & SVM training and honest evaluation
7. Classification map generation
"""

import sys
from pathlib import Path

# Add project root to sys.path
PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import numpy as np
import pandas as pd
import geopandas as gpd
import ee

from config import (
    LABELS_DIR,
    PROCESSED_DATA_DIR,
    OUTPUTS_DIR,
    GEE_PROJECT_ID,
    TIME_STEPS,
    TARGET_CROPS,
    CROP_COLORS,
    CROP_MAP_PNG,
    CROP_MAP_TIF
)
from src.gee_utils import initialize_gee, build_seasonal_composites
from src.dataset import (
    find_label_files,
    load_and_filter_labels,
    buffer_field_polygons,
    spatial_train_test_split,
    generate_dataset_summary
)
from src.features import compute_spectral_indices_np, plot_ndvi_phenology
from src.models import train_and_evaluate_models
from src.mapping import export_classification_map


def extract_field_features(gdf: gpd.GeoDataFrame, chunk_size: int = 100) -> pd.DataFrame:
    """
    Extract median spectral reflectance per field polygon across all seasonal composites
    using fast server-side batch Earth Engine reduceRegions chunked by local spatial bounds.
    """
    import shapely.geometry
    from src.gee_utils import get_s2_composite
    print(f"[*] Extracting multi-temporal Sentinel-2 features for {len(gdf)} fields across {len(TIME_STEPS)} time steps...")
    
    all_rows = []
    num_chunks = int(np.ceil(len(gdf) / chunk_size))
    
    for c_idx in range(num_chunks):
        chunk_gdf = gdf.iloc[c_idx * chunk_size : (c_idx + 1) * chunk_size]
        minx, miny, maxx, maxy = [float(x) for x in chunk_gdf.total_bounds]
        chunk_roi = ee.Geometry.Rectangle([minx - 0.05, miny - 0.05, maxx + 0.05, maxy + 0.05])
        
        print(f"    -> Querying GEE batch {c_idx + 1}/{num_chunks} ({len(chunk_gdf)} fields)...", flush=True)
        
        composites = {label: get_s2_composite(chunk_roi, s, e) for s, e, label in TIME_STEPS}
        stacked_image = ee.Image.cat([
            comp.select(["B2", "B3", "B4", "B8", "B11"], [f"{lbl}_B2", f"{lbl}_B3", f"{lbl}_B4", f"{lbl}_B8", f"{lbl}_B11"])
            for lbl, comp in composites.items()
        ])
        
        features = []
        for idx, row in chunk_gdf.iterrows():
            geom_json = shapely.geometry.mapping(row.geometry)
            props = {
                "field_id": int(row["field_id"]),
                "crop_name": str(row["crop_name"]),
                "area_ha": float(row.get("area_ha", 0.0)),
                "tile_id": str(row.get("tile_id", f"tile_{idx}"))
            }
            features.append(ee.Feature(ee.Geometry(geom_json), props))
            
        fc = ee.FeatureCollection(features)
        reduced_fc = stacked_image.reduceRegions(
            collection=fc,
            reducer=ee.Reducer.median(),
            scale=10
        )
        
        results = reduced_fc.getInfo()
        for feat in results["features"]:
            props = feat["properties"]
            all_rows.append(props)

    df = pd.DataFrame(all_rows)
    df = df.dropna().reset_index(drop=True)
    print(f"[+] Successfully extracted complete time-series features for {len(df)} fields.")
    return df


def run_pipeline():
    print("=" * 70)
    print("Crop Classification End-to-End Pipeline (BIT Mesra Minor Project)")
    print("=" * 70)

    # 1. Check GEE Configuration
    print("\n[Stage 1 & 2] Checking Earth Engine & Sentinel-2...")
    try:
        initialize_gee()
        print(f"[+] Earth Engine initialized with project: {GEE_PROJECT_ID}")
    except Exception as e:
        print(f"[-] GEE Initialization halted: {e}")
        print("    Please follow Stage 1 instructions in README.md to authenticate.")
        return

    # 2. Check Ground Truth Labels
    print("\n[Stage 4] Searching for AgriFieldNet field polygons...")
    label_files = find_label_files(LABELS_DIR)
    if not label_files:
        print(f"[!] No label files found in {LABELS_DIR}.")
        print("    Please place your AgriFieldNet India field polygons (.geojson, .shp, or .csv)")
        print(f"    in: {LABELS_DIR}")
        return

    label_path = label_files[0]
    print(f"[+] Found label file: {label_path}")
    
    # Load, filter 3 crops, buffer
    gdf = load_and_filter_labels(label_path)
    gdf_buffered = buffer_field_polygons(gdf)
    train_gdf, test_gdf = spatial_train_test_split(gdf_buffered)
    
    # Save dataset summary table for Chapter 3 of report
    summary_df = generate_dataset_summary(gdf_buffered, train_gdf, test_gdf)

    # 3. Extract Features via Cloud Score+ Sentinel-2 Composites
    features_csv = PROCESSED_DATA_DIR / "field_features_all.csv"
    time_labels = [step[2] for step in TIME_STEPS]
    
    if features_csv.exists() and len(pd.read_csv(features_csv)) == len(gdf_buffered):
        print(f"[*] Loading cached multi-temporal features for {len(gdf_buffered)} fields...")
        features_df = pd.read_csv(features_csv)
    else:
        print("\n[Stage 2 & 3] Extracting multi-temporal Sentinel-2 features (Cloud Score+ QA)...")
        raw_features_df = extract_field_features(gdf_buffered)
        features_df = compute_spectral_indices_np(raw_features_df, time_labels)
        features_df.to_csv(features_csv, index=False)
        print(f"[+] Saved extracted feature matrix to {features_csv}")
    
    # Plot phenology trajectories
    print("\n[Stage 3] Plotting NDVI phenology curves...")
    plot_ndvi_phenology(features_df)

    # 4. Prepare Train / Test Splits for ML
    feature_cols = []
    for t in time_labels:
        for idx_name in ["NDVI", "EVI", "NDWI"]:
            feature_cols.append(f"{t}_{idx_name}")

    train_df = features_df[features_df["field_id"].isin(train_gdf["field_id"])]
    test_df = features_df[features_df["field_id"].isin(test_gdf["field_id"])]

    X_train = train_df[feature_cols]
    y_train = train_df["crop_name"]
    X_test = test_df[feature_cols]
    y_test = test_df["crop_name"]

    # 5. Train & Evaluate Models (Random Forest baseline & SVM)
    print("\n[Stage 5] Training Random Forest and SVM with spatial evaluation...")
    eval_results = train_and_evaluate_models(
        X_train=X_train,
        y_train=y_train,
        X_test=X_test,
        y_test=y_test,
        feature_names=feature_cols
    )

    # 6. Generate Crop Classification Map
    print("\n[Stage 6] Exporting Crop Classification Map...")
    rf_model = eval_results["models"]["Random Forest"]
    features_df["predicted_crop"] = rf_model.predict(features_df[feature_cols])
    
    # Merge predictions with georeferenced field polygons
    classified_gdf = gdf_buffered.merge(
        features_df[["field_id", "predicted_crop"]],
        on="field_id",
        how="inner"
    )
    # The Bahraich-Gonda agricultural belt in Uttar Pradesh
    core_bounds = (82.15, 27.35, 82.75, 27.65)
    export_classification_map(classified_gdf, bounds=core_bounds)

    print("\n" + "=" * 70)
    print("STAGES 1 - 7 COMPLETE!")
    print(f"Metrics: {OUTPUTS_DIR / 'metrics.json'}")
    print(f"Figures: {OUTPUTS_DIR}")
    print("Ready for user review and Streamlit app launch.")
    print("=" * 70)


if __name__ == "__main__":
    run_pipeline()
