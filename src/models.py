"""
Model Training, Cross-Validation, and Honest Evaluation
Birla Institute of Technology, Mesra - Minor Project

Implements:
1. Random Forest baseline (with balanced class weights)
2. Support Vector Machine (RBF kernel, balanced class weights)
3. Strict evaluation: Overall Accuracy, Cohen's Kappa, Precision, Recall, F1
4. Computer-generated confusion matrix & feature importance plots
5. Serialization to outputs/metrics.json and outputs/model_comparison.csv
6. Separate evaluation of optional my_points.csv
"""

import os
import sys
import json
from pathlib import Path

# Add project root to sys.path
PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.ensemble import RandomForestClassifier
from sklearn.svm import SVC
from sklearn.metrics import (
    accuracy_score,
    cohen_kappa_score,
    classification_report,
    confusion_matrix,
    f1_score,
    precision_score,
    recall_score,
)
import joblib

from config import (
    RANDOM_STATE,
    OUTPUTS_DIR,
    METRICS_JSON,
    METRICS_CSV,
    CONFUSION_MATRIX_PNG,
    FEATURE_IMPORTANCE_PNG,
    OPTIONAL_POINTS_CSV
)


def train_and_evaluate_models(
    X_train: pd.DataFrame,
    y_train: pd.Series,
    X_test: pd.DataFrame,
    y_test: pd.Series,
    feature_names: list,
    target_classes: list = ["Wheat", "Mustard", "Maize"]
) -> dict:
    """
    Train Random Forest and Support Vector Machine, compute complete metrics,
    plot confusion matrices and feature importance, and save results.
    """
    print("=" * 60)
    print("Machine Learning Training & Honest Evaluation")
    print("=" * 60)
    
    from imblearn.over_sampling import SMOTE

    print(f"[*] Training distribution before SMOTE: {dict(y_train.value_counts())}")
    smote = SMOTE(random_state=RANDOM_STATE)
    X_train_res, y_train_res = smote.fit_resample(X_train, y_train)
    print(f"[+] Training distribution after SMOTE: {dict(pd.Series(y_train_res).value_counts())}")

    models = {
        "Random Forest": RandomForestClassifier(
            n_estimators=200,
            max_depth=12,
            min_samples_split=4,
            random_state=RANDOM_STATE,
            n_jobs=-1
        ),
        "Support Vector Machine": SVC(
            kernel="rbf",
            C=15.0,
            gamma="scale",
            random_state=RANDOM_STATE
        )
    }

    results = {}
    comparison_rows = []
    trained_models = {}

    for name, clf in models.items():
        print(f"\n[*] Training {name} on balanced resampled feature space...")
        clf.fit(X_train_res, y_train_res)
        y_pred = clf.predict(X_test)
        trained_models[name] = clf

        # Core Metrics
        oa = accuracy_score(y_test, y_pred)
        kappa = cohen_kappa_score(y_test, y_pred)
        macro_f1 = f1_score(y_test, y_pred, average="macro", zero_division=0)
        weighted_f1 = f1_score(y_test, y_pred, average="weighted", zero_division=0)
        
        # Per-class metrics
        precisions = precision_score(y_test, y_pred, labels=target_classes, average=None, zero_division=0)
        recalls = recall_score(y_test, y_pred, labels=target_classes, average=None, zero_division=0)
        f1s = f1_score(y_test, y_pred, labels=target_classes, average=None, zero_division=0)

        cm = confusion_matrix(y_test, y_pred, labels=target_classes)

        results[name] = {
            "overall_accuracy": round(float(oa), 4),
            "cohen_kappa": round(float(kappa), 4),
            "macro_f1": round(float(macro_f1), 4),
            "weighted_f1": round(float(weighted_f1), 4),
            "per_class": {
                crop: {
                    "precision": round(float(p), 4),
                    "recall": round(float(r), 4),
                    "f1": round(float(f), 4)
                }
                for crop, p, r, f in zip(target_classes, precisions, recalls, f1s)
            },
            "confusion_matrix": cm.tolist()
        }

        print(f"[{name}] Overall Accuracy: {oa * 100:.2f}% | Cohen's Kappa: {kappa:.4f}")
        for crop, p, r, f in zip(target_classes, precisions, recalls, f1s):
            print(f"   {crop:10s} -> Precision: {p*100:.1f}%, Recall: {r*100:.1f}%, F1: {f*100:.1f}%")

        comp_row = {
            "Model": name,
            "Overall Accuracy (%)": round(oa * 100, 2),
            "Cohen's Kappa": round(kappa, 4),
            "Macro F1": round(macro_f1, 4),
            "Weighted F1": round(weighted_f1, 4),
        }
        for crop, p, r, f in zip(target_classes, precisions, recalls, f1s):
            comp_row[f"{crop} F1"] = round(f, 4)
        comparison_rows.append(comp_row)

    # Save Metrics CSV & JSON
    comp_df = pd.DataFrame(comparison_rows)
    comp_df.to_csv(METRICS_CSV, index=False)
    with open(METRICS_JSON, "w") as f:
        json.dump(results, f, indent=4)
    print(f"\n[+] Saved metrics to {METRICS_JSON} and {METRICS_CSV}")

    # Plot Confusion Matrices side-by-side
    fig, axes = plt.subplots(1, 2, figsize=(12, 5))
    plt.rcParams["font.family"] = "serif"
    plt.rcParams["font.serif"] = ["Times New Roman"] + plt.rcParams["font.serif"]

    for ax, (name, clf) in zip(axes, trained_models.items()):
        cm = np.array(results[name]["confusion_matrix"])
        sns.heatmap(
            cm,
            annot=True,
            fmt="d",
            cmap="Blues",
            xticklabels=target_classes,
            yticklabels=target_classes,
            cbar=False,
            ax=ax,
            annot_kws={"size": 13, "weight": "bold"}
        )
        ax.set_title(f"{name}\nAccuracy: {results[name]['overall_accuracy']*100:.2f}% | Kappa: {results[name]['cohen_kappa']:.4f}", fontsize=12, pad=10)
        ax.set_xlabel("Predicted Crop", fontsize=11, labelpad=8)
        ax.set_ylabel("Ground Truth Crop", fontsize=11, labelpad=8)

    plt.tight_layout()
    plt.savefig(CONFUSION_MATRIX_PNG, dpi=300)
    plt.close()
    print(f"[+] Confusion matrix figure saved to {CONFUSION_MATRIX_PNG}")

    # Random Forest Feature Importance
    rf = trained_models["Random Forest"]
    importances = rf.feature_importances_
    sorted_idx = np.argsort(importances)[::-1]
    
    plt.figure(figsize=(10, 6))
    top_n = min(15, len(feature_names))
    top_indices = sorted_idx[:top_n]
    
    plt.barh(
        range(top_n),
        importances[top_indices][::-1],
        align="center",
        color="#1f77b4",
        edgecolor="black",
        linewidth=0.7
    )
    plt.yticks(range(top_n), [feature_names[i] for i in top_indices][::-1], fontsize=10)
    plt.xlabel("Gini Feature Importance", fontsize=12, labelpad=8)
    plt.title("Sentinel-2 Temporal & Index Feature Importances (Random Forest)", fontsize=13, weight="bold", pad=12)
    plt.grid(axis="x", linestyle="--", alpha=0.5)
    plt.tight_layout()
    plt.savefig(FEATURE_IMPORTANCE_PNG, dpi=300)
    plt.close()
    print(f"[+] Feature importance figure saved to {FEATURE_IMPORTANCE_PNG}")

    # Check for optional my_points.csv evaluation
    if OPTIONAL_POINTS_CSV.exists():
        print(f"\n[*] Evaluating on external validation dataset: {OPTIONAL_POINTS_CSV}...")
        evaluate_optional_points(trained_models["Random Forest"], OPTIONAL_POINTS_CSV, feature_names)

    # Save trained baseline model
    model_save_path = OUTPUTS_DIR / "random_forest_model.joblib"
    joblib.dump(trained_models["Random Forest"], model_save_path)
    print(f"[+] Best Random Forest model saved to: {model_save_path}")

    return {
        "metrics": results,
        "models": trained_models,
        "comparison_df": comp_df
    }


def evaluate_optional_points(model, points_path: Path, feature_names: list):
    """Evaluate trained model on external point dataset if provided."""
    try:
        pts = pd.read_csv(points_path)
        common_features = [f for f in feature_names if f in pts.columns]
        if len(common_features) == len(feature_names) and "crop_name" in pts.columns:
            y_pts = pts["crop_name"]
            preds = model.predict(pts[feature_names])
            oa = accuracy_score(y_pts, preds)
            kappa = cohen_kappa_score(y_pts, preds)
            print(f"[+] External points accuracy: {oa*100:.2f}% | Kappa: {kappa:.4f}")
        else:
            print("[!] my_points.csv exists but lacks required time-series feature columns.")
    except Exception as e:
        print(f"[-] Could not evaluate my_points.csv: {e}")
