import os
import requests

DATA_URL = "https://raw.githubusercontent.com/rahulshirke7951/amfi-analytics-engine/main/output/dashboard_data.xlsx"
SAVE_PATH = "dashboard_data.xlsx"   # IMPORTANT: same as your main script

def download_file():
    print("📥 Checking data file...")

    if os.path.exists(SAVE_PATH):
        print("✅ File already exists. Skipping download.")
        return

    print("⬇️ Downloading dashboard_data.xlsx...")

    response = requests.get(DATA_URL)

    if response.status_code == 200:
        with open(SAVE_PATH, "wb") as f:
            f.write(response.content)
        print("✅ Download complete.")
    else:
        raise Exception(f"❌ Failed to download file. Status: {response.status_code}")

if __name__ == "__main__":
    download_file()
