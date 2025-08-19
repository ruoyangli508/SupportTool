import os
import time
import requests
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from tkinter import filedialog, Tk
from concurrent.futures import ThreadPoolExecutor, as_completed

API_URL = "https://trk.speedx.io/tracking-api/pod/listLabelFile"


def choose_file():
    root = Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename()
    if not file_path:
        raise FileNotFoundError("No file selected")
    return file_path


def call_api_with_retry(trk_list, max_retries=3, delay=2):
    """
    Call POD API with retry mechanism
    """
    for attempt in range(max_retries):
        try:
            resp = requests.post(
                API_URL,
                headers={"Content-Type": "application/json"},
                json={"fileType": "", "trackingNumbers": trk_list},
                timeout=30
            )
            if resp.status_code == 200:
                data = resp.json()
                if data.get("success"):
                    return data.get("payload", [])
            print(f"API call failed, status {resp.status_code}, retry {attempt+1}")
        except Exception as e:
            print(f"Request exception: {e}, retry {attempt+1}")

        time.sleep(delay)
    return []


def download_file(item, base_dir):
    """
    Download a single POD file into its trackingNumber folder
    """
    tracking_number = item["trackingNumber"]
    file_url = item["fileUrl"]
    ext = os.path.splitext(file_url)[1] or ".jpg"

    trk_dir = os.path.join(base_dir, tracking_number)
    os.makedirs(trk_dir, exist_ok=True)

    save_path = os.path.join(trk_dir, os.path.basename(file_url))
    try:
        r = requests.get(file_url, timeout=30)
        if r.status_code == 200:
            with open(save_path, "wb") as f:
                f.write(r.content)
    except Exception as e:
        # No print (silent fail as required)
        pass


def get_pod_data(all_trk_list, save_dir):
    """
    Fetch POD data in batches (10 per request) and download files with 4 threads
    """
    results = []
    total_batches = (len(all_trk_list) + 9) // 10

    for i in range(0, len(all_trk_list), 10):
        batch_num = i // 10 + 1
        print(f"Processing batch {batch_num}/{total_batches} ...")

        batch = all_trk_list[i:i + 10]
        payload = call_api_with_retry(batch)
        if not payload:
            continue

        results.extend(payload)

        # Parallel download with 4 threads
        with ThreadPoolExecutor(max_workers=4) as executor:
            futures = [executor.submit(download_file, item, save_dir) for item in payload]
            for _ in as_completed(futures):
                pass

    return results


def write_to_excel(pod_data, source_df, save_path):
    """
    Write results into Excel:
    - sheet1: pod_data
    - sheet2: uploaded_tracking_number
    """
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "pod_data"

    if pod_data:
        df_pod = pd.DataFrame(pod_data)
        for r in dataframe_to_rows(df_pod, index=False, header=True):
            ws1.append(r)

    ws2 = wb.create_sheet("uploaded_tracking_number")
    for r in dataframe_to_rows(source_df, index=False, header=True):
        ws2.append(r)

    wb.save(save_path)


def main():
    print("Get POD Tool")
    print("Make sure that the tracking numbers are all in the first column")
    print("\n")
    print("------------------------------------------------------")
    input("Press Enter to start (≧▽≦) ")

    try:
        file_path = choose_file()
        df = pd.read_excel(file_path)
        trk_list = list(set(df.iloc[:, 0].dropna().astype(str)))

        base_dir = os.path.dirname(file_path)
        file_name = os.path.splitext(os.path.basename(file_path))[0]

        # Create new folder for images
        pod_folder = os.path.join(base_dir, file_name)
        os.makedirs(pod_folder, exist_ok=True)

        # Excel report path
        report_file = os.path.join(base_dir, f"{file_name}_pod_result.xlsx")

        # Fetch data & download files
        pod_data = get_pod_data(trk_list, pod_folder)

        # Write Excel report
        write_to_excel(pod_data, df, report_file)

        print(f"All done! Report saved to: {report_file}")
        print(f"All POD images saved in folder: {pod_folder}")

    except Exception as e:
        print(f"An unexpected error occurred: {e}")

    input("Press Enter to exit (o´ω`o)ﾉ ")




if __name__ == "__main__":
    main()
