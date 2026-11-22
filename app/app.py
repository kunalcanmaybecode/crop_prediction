"""
Streamlit Web Dashboard for Crop Classification System
Birla Institute of Technology, Mesra - Minor Project

Features:
- Regional classified crop map with interactive layer toggle
- Per-crop NDVI phenological trajectory analysis
- Real model accuracy metrics, Cohen's Kappa, and per-class metrics
- Computer-generated confusion matrix & feature importance plots
- Dataset distribution and field area statistics
"""

import json
from pathlib import Path
import streamlit as st
import pandas as pd
import numpy as np
from PIL import Image

# Page configuration
st.set_page_config(
    page_title="Crop Classification System | BIT Mesra",
    page_icon="🌾",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Paths
BASE_DIR = Path(__file__).resolve().parent.parent
OUTPUTS_DIR = BASE_DIR / "outputs"
METRICS_JSON = OUTPUTS_DIR / "metrics.json"
METRICS_CSV = OUTPUTS_DIR / "model_comparison.csv"
DATASET_STATS_CSV = OUTPUTS_DIR / "dataset_summary.csv"
CONFUSION_MATRIX_PNG = OUTPUTS_DIR / "confusion_matrix.png"
FEATURE_IMPORTANCE_PNG = OUTPUTS_DIR / "feature_importance.png"
NDVI_PHENOLOGY_PNG = OUTPUTS_DIR / "crop_phenology_ndvi.png"
CROP_MAP_PNG = OUTPUTS_DIR / "crop_classification_map.png"
CROP_MAP_TIF = OUTPUTS_DIR / "crop_classification_map.tif"

# Header
st.title("🌾 Multi-Temporal Crop Classification System")
st.markdown("*Sentinel-2 Harmonized Surface Reflectance & Machine Learning Dashboard*")
st.divider()

# Sidebar Navigation
st.sidebar.title("Navigation")
section = st.sidebar.radio(
    "Select View:",
    ["Overview & Dataset", "Model Performance & Metrics", "Phenology Curves (NDVI)", "Crop Classification Map"]
)

st.sidebar.info(
    "**Target Crops:**\n"
    "- 🌾 **Wheat** (Rabi cereal)\n"
    "- 🌼 **Mustard** (Rabi oilseed)\n"
    "- 🌽 **Maize** (Coarse grain)\n\n"
    "**Study Area:** Uttar Pradesh, India\n"
    "**Sensor:** Sentinel-2 (MSI)\n"
    "**Cloud Mask:** Cloud Score+"
)

# Section 1: Overview & Dataset
if section == "Overview & Dataset":
    st.header("📊 Dataset Overview (AgriFieldNet India)")
    st.markdown(
        """
        Ground truth field boundaries from **AgriFieldNet India** were filtered strictly to the three target crops:
        **Wheat**, **Mustard**, and **Maize**. A **5-meter inward buffer** was applied to each field polygon to 
        eliminate mixed boundary pixels. A **spatial train/test split** was enforced to prevent spatial autocorrelation data leakage.
        """
    )
    
    col1, col2, col3 = st.columns(3)
    if DATASET_STATS_CSV.exists():
        df_summary = pd.read_csv(DATASET_STATS_CSV)
        total_fields = df_summary["Total Fields"].sum()
        total_area = df_summary["Total Area (ha)"].sum()
        col1.metric("Total Valid Fields", f"{total_fields:,}")
        col2.metric("Total Mapped Area", f"{total_area:,.1f} ha")
        col3.metric("Spatial Split", "75% Train / 25% Test")
        
        st.subheader("Field Distribution and Area Statistics")
        st.dataframe(df_summary, use_container_width=True)
    else:
        st.warning("Dataset summary not generated yet. Run `python src/pipeline.py` after configuring GEE.")
        col1.metric("Classes", "3 Crops")
        col2.metric("Sensor", "Sentinel-2")
        col3.metric("Composites", "6 Monthly")

# Section 2: Model Performance & Metrics
elif section == "Model Performance & Metrics":
    st.header("📈 Model Evaluation & Comparison")
    st.markdown(
        """
        Comparison between **Random Forest** (baseline) and **Support Vector Machine (SVM)**.
        All metrics reflect genuine evaluation on the unseen spatial test partition.
        """
    )

    if METRICS_JSON.exists():
        with open(METRICS_JSON, "r") as f:
            metrics = json.load(f)

        rf = metrics.get("Random Forest", {})
        svm = metrics.get("Support Vector Machine", {})

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("RF Overall Accuracy", f"{rf.get('overall_accuracy', 0)*100:.2f}%")
        c2.metric("RF Cohen's Kappa", f"{rf.get('cohen_kappa', 0):.4f}")
        c3.metric("SVM Overall Accuracy", f"{svm.get('overall_accuracy', 0)*100:.2f}%")
        c4.metric("SVM Cohen's Kappa", f"{svm.get('cohen_kappa', 0):.4f}")

        if METRICS_CSV.exists():
            st.subheader("Summary Metric Comparison Table")
            df_comp = pd.read_csv(METRICS_CSV)
            st.dataframe(df_comp, use_container_width=True)

        st.subheader("Confusion Matrix")
        if CONFUSION_MATRIX_PNG.exists():
            st.image(str(CONFUSION_MATRIX_PNG), caption="Confusion Matrices on Spatial Test Partition", use_container_width=True)

        st.subheader("Feature Importance (Random Forest)")
        if FEATURE_IMPORTANCE_PNG.exists():
            st.image(str(FEATURE_IMPORTANCE_PNG), caption="Top Temporal Indices Influencing Crop Discrimination", use_container_width=True)
    else:
        st.warning("Metrics file not found. Please run the model training pipeline first.")

# Section 3: Phenology Curves (NDVI)
elif section == "Phenology Curves (NDVI)":
    st.header("🌱 Crop Phenology (Temporal NDVI Curves)")
    st.markdown(
        """
        Multi-temporal vegetation curves capture the distinct life stages (phenology) of crops across the Rabi season:
        - **Mustard:** Rapid early vegetative growth in Nov-Dec, bright yellow flowering canopy in Dec-Jan, and early harvest in Feb.
        - **Wheat:** Gradual tillering in Dec, maximum green canopy/heading in Feb, followed by dry-down/senescence in Mar-Apr.
        - **Maize:** Distinct emergence and biomass curve differing in peak timing and water index absorption.
        """
    )
    if NDVI_PHENOLOGY_PNG.exists():
        st.image(str(NDVI_PHENOLOGY_PNG), caption="Mean Sentinel-2 NDVI Phenological Trajectories with SEM Bands", use_container_width=True)
    else:
        st.info("Phenology curves will appear here once features are extracted.")

# Section 4: Crop Classification Map
elif section == "Crop Classification Map":
    st.header("🗺️ Regional Crop Classification Map")
    st.markdown(
        """
        Cartographic classification map over the Uttar Pradesh study region.
        Each field pixel is assigned to one of the 3 target crops based on its seasonal spectral signature.
        """
    )
    
    col_a, col_b = st.columns([3, 1])
    with col_a:
        if CROP_MAP_PNG.exists():
            st.image(str(CROP_MAP_PNG), caption="Classified Crop Map (Green: Wheat, Pink/Yellow: Mustard, Orange: Maize)", use_container_width=True)
        else:
            st.info("Classification map will be rendered here upon running the mapping stage.")

    with col_b:
        st.subheader("Map Legend")
        st.markdown(
            """
            - 🟩 **Wheat** (Code 1)
            - 🟨 **Mustard** (Code 2)
            - 🟧 **Maize** (Code 3)
            - ⬜ **Non-target / Background**
            """
        )
        if CROP_MAP_TIF.exists():
            with open(CROP_MAP_TIF, "rb") as f:
                st.download_button("📥 Download GeoTIFF", f, file_name="crop_classification_map.tif", mime="image/tiff")

st.divider()
