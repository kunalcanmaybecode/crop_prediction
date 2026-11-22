"""
Verification Script for Google Earth Engine (GEE)
Birla Institute of Technology, Mesra - Minor Project

Tests:
1. GEE library import
2. Authentication & project initialization
3. Basic test query (Sentinel-2 Harmonized collection over UP)
"""

import sys
import ee
from config import GEE_PROJECT_ID, S2_COLLECTION, CLOUD_SCORE_COLLECTION

def test_earth_engine():
    print("=" * 60)
    print("Google Earth Engine (GEE) Verification Test")
    print("=" * 60)
    
    if GEE_PROJECT_ID == "<MY_GEE_PROJECT_ID>" or not GEE_PROJECT_ID:
        print("[!] GEE_PROJECT_ID is not configured yet in config.py or environment.")
        print("    Please follow the instructions in README.md to obtain your GEE Project ID.")
        print("    You can run:")
        print("      $env:GEE_PROJECT_ID='your-cloud-project-id'")
        print("      python test_gee.py")
        return False

    print(f"[*] Attempting to initialize Earth Engine with project: '{GEE_PROJECT_ID}'...")
    try:
        ee.Initialize(project=GEE_PROJECT_ID)
        print("[+] Success: Earth Engine initialized successfully!")
    except Exception as e:
        print(f"[-] Initialization failed: {e}")
        print("\nTroubleshooting steps:")
        print("1. If you haven't authenticated on this machine, run:")
        print("   earthengine authenticate")
        print("2. Ensure the Google Cloud Project has Earth Engine API enabled:")
        print("   Visit: https://console.cloud.google.com/apis/library/earthengine.googleapis.com")
        print("3. Ensure your Google account is registered as an Earth Engine user:")
        print("   Visit: https://code.earthengine.google.com/register")
        return False

    print("\n[*] Running a tiny test query...")
    try:
        # Test Point in Uttar Pradesh (approx. near Lucknow / Central UP)
        up_point = ee.Geometry.Point([80.9462, 26.8467])
        
        # Test query on Sentinel-2 Harmonized collection
        s2_test = (
            ee.ImageCollection(S2_COLLECTION)
            .filterBounds(up_point)
            .filterDate("2022-01-01", "2022-01-15")
        )
        count = s2_test.size().getInfo()
        print(f"[+] Query successful! Found {count} Sentinel-2 scenes over test point in Jan 2022.")

        # Test Cloud Score+ collection availability
        cs_test = (
            ee.ImageCollection(CLOUD_SCORE_COLLECTION)
            .filterBounds(up_point)
            .filterDate("2022-01-01", "2022-01-15")
        )
        cs_count = cs_test.size().getInfo()
        print(f"[+] Cloud Score+ dataset verified! Found {cs_count} matching QA scenes.")
        print("\n>>> ALL GEE CHECKS PASSED SUCCESSFULLY! Ready for Stage 2. <<<")
        return True
    except Exception as e:
        print(f"[-] Test query failed: {e}")
        return False

if __name__ == "__main__":
    success = test_earth_engine()
    sys.exit(0 if success else 1)
