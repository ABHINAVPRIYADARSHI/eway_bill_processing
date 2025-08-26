import os, sys
import win32com.client as win32
import gc
from io import StringIO
import pandas as pd
import numpy as np
import json
import time, calendar
from datetime import datetime
from glob import glob
from playwright.sync_api import sync_playwright, Page
from pathlib import Path

# Config and log paths passed as arguments
CONFIG_PATH = sys.argv[1]
LOG_PATH = sys.argv[2]
# Mapping of month names to numbers
month_name_to_number = {
    "January": 1, "February": 2, "March": 3,
    "April": 4, "May": 5, "June": 6,
    "July": 7, "August": 8, "September": 9,
    "October": 10, "November": 11, "December": 12
}
in_radio_button = 'input[id="ctl00_ContentPlaceHolder1_RBL_OutInward_1"]'
out_radio_button = 'input[id="ctl00_ContentPlaceHolder1_RBL_OutInward_0"]'
EWB_MIS_Report_Excel = 'EWB_MIS_Report_Excel'
DEFAULT_TIMEOUT = 180000 # 180 sec or 3 mins
_5_MIN_TIMEOUT = 300000 # 300 sec or 5 mins
_IN_ = "In"
_OUT_ = "Out"
# Map of state option values to state group names
state_options = {
    "0": "Select State",  # Skip this default option
    "1": "Andhra_Pradesh_Goa_Karnataka_Telangana",
    "2": "Maharashtra", 
    "3": "Chandigarh_Haryana_HimachalPradesh_JammuKashmir_Punjab_Uttarakhand",
    "4": "Jharkhand_UttarPradesh",
    "5": "Kerala_Lakshadweep_Puducherry_TamilNadu",
    "6": "DadraNagarHaveli_Daman_Diu_Gujarat_MadhyaPradesh_Chhattisgarh",
    "7": "Delhi_Rajasthan"
}

def log(msg: str):
    timestamp = time.strftime('%Y-%m-%d %H:%M:%S')
    with open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write(f"{timestamp} - {msg}\n")
    print(f"{timestamp} - {msg}")


def get_days_in_month(month_year: tuple):
    month_name, year = month_year
    month_num = month_name_to_number[month_name]
    today = datetime.today() # Get today's date

    # Check if current month/year
    if today.month == month_num and today.year == year:
        return today.day  # Return today's date
    # Otherwise, return number of days in that month/year
    return calendar.monthrange(year, month_num)[1]


def login_and_open_ewb_mis(page: Page, context, username: str, password: str) -> Page:
    log("Opening EWB login page...")
    page.goto("https://gstsso.nic.in/")
    page.locator("#txt_username").fill(username)
    page.locator("#txt_password").fill(password)

    log("Waiting for manual CAPTCHA + OTP entry...")
    page.wait_for_url("**/webfrmdd.aspx", timeout=_5_MIN_TIMEOUT)
    log("✅ Successfully logged in to GST portal.")
    page.wait_for_load_state("networkidle", timeout=_5_MIN_TIMEOUT)
    log("Clicking button to open EWB MIS portal...")
    # Wait for the btnewbmis button to be available
    button_selector = 'input[name="btnewbmis"]'
    page.wait_for_selector(button_selector, timeout=_5_MIN_TIMEOUT)
    log("EWB MIS button found, clicking...")
    
    # Set up listener for new tab BEFORE clicking
    with context.expect_page(timeout=_5_MIN_TIMEOUT) as new_page_info:
        page.click(button_selector)
        log("EWB MIS button clicked, waiting for new tab...")
    
    # Get the new page (EWB MIS portal)
    ewb_mis_page = new_page_info.value
    
    # Wait for the new page to fully load
    ewb_mis_page.wait_for_load_state('networkidle', timeout=_5_MIN_TIMEOUT)
    log(f"EWB MIS portal opened successfully: {ewb_mis_page.url}")
    
    ewb_mis_page.goto("https://mis.ewaybillgst.gov.in/Verification/GSTINBasedRpt.aspx", timeout=_5_MIN_TIMEOUT)
    # Optional: wait for DOM or network idle
    # ewb_mis_page.wait_for_load_state("domcontentloaded", timeout=_5_MIN_TIMEOUT)
    ewb_mis_page.wait_for_load_state("networkidle", timeout=_5_MIN_TIMEOUT)
    log(f"✅ GSTINBasedRpt.aspx portal opened successfully: URL= {ewb_mis_page.url}")
    # Bring the new tab to front
    ewb_mis_page.bring_to_front()
    log("EWB MIS portal is ready for use.")
    return ewb_mis_page


def get_month_year_range(start_month: str, start_year, end_month: str, end_year):
    result = []
    try:
        # Convert month strings to integers
        start_month_num = month_name_to_number[start_month]
        end_month_num = month_name_to_number[end_month]
        # Reverse lookup to get full month name from number
        number_to_month_name = {v: k for k, v in month_name_to_number.items()}
        # Start loop from start year/month to end year/month
        year = start_year
        month = start_month_num
        while (year < end_year) or (year == end_year and month <= end_month_num):
            full_month_name = number_to_month_name[month]
            # short_month = full_month_name[:3]
            result.append((full_month_name, year))
            # Increment month/year
            if month == 12:
                month = 1
                year += 1
            else:
                month += 1
    except Exception as e:
        log(f"❌ Error while calculating month_year_range: {e}")
    return result

# finally:
    # Reset page
    # try:
    #     # page.goto("https://mis.ewaybillgst.gov.in/Verification/GSTINBasedRpt.aspx")
    #     # page.wait_for_selector('xpath=//*[@id="ctl00_ContentPlaceHolder1_txt_gstin"]', timeout=15000)
    # except Exception as e:
    #     log(f"Failed to reset page after {statecombo}: {e}")


def _get_month_number(month_name: str) -> int:
    """Convert month name to number."""
    return month_name_to_number.get(month_name, 1)


def _get_radio_button_selector(radio_type: str) -> str:
    """Get the appropriate radio button selector based on type."""
    if radio_type == _OUT_:
        return out_radio_button
    elif radio_type == _IN_:
        return in_radio_button
    else:
        raise ValueError(f"❌ Invalid radio button type: {radio_type}")


def _set_date_fields_exact(page: Page, last_day: int, month_year: tuple):
    """
    Set the date fields using exact field IDs with JavaScript.
    """
    month, year = month_year
    from_date = f"01/{_get_month_number(month):02d}/{year}"
    to_date = f"{last_day:02d}/{_get_month_number(month):02d}/{year}"
    
    # Use JavaScript to set readonly date fields
    page.wait_for_selector('#ctl00_ContentPlaceHolder1_txtDateFrom', timeout= DEFAULT_TIMEOUT)
    page.evaluate(f'document.getElementById("ctl00_ContentPlaceHolder1_txtDateFrom").value = "{from_date}"')
    page.wait_for_selector('#ctl00_ContentPlaceHolder1_txtDateTo', timeout=DEFAULT_TIMEOUT)
    page.evaluate(f'document.getElementById("ctl00_ContentPlaceHolder1_txtDateTo").value = "{to_date}"')
    log(f"Set date range: {from_date} to {to_date}")


def _check_for_export_to_excel(page):
    """
    Check if data results are available on the page by looking for the Export to Excel button
    """
    try:
        # Method 1: Check for the Export to Excel button by ID. We don't need to give timeout here
        # beacuse we have already waited for "networkidle" in previous method.
        page.wait_for_selector("#ctl00_ContentPlaceHolder1_btn_export_excel", timeout=100)
        return True
    except Exception as e:
        return False
    

def _click_go_and_download_excel(page: Page, gstin: str, state_name: str, month_year: tuple, downloads_dir: str, in_out_prefix):
    """
    Click GO button, check for data, and download Excel if available.
    Returns:
        bool: True if Excel was downloaded, False otherwise
    """
    try:
        # Click GO button
        go_button = 'input[name="ctl00$ContentPlaceHolder1$btnsbmt"][value="GO"]'
        page.wait_for_selector(go_button, timeout=_5_MIN_TIMEOUT)
        page.click(go_button)
        # Wait for results to load (lighter wait)
        page.wait_for_load_state("networkidle", timeout= DEFAULT_TIMEOUT)
        # Check if any data is available before setting up 
        file_name = f"{in_out_prefix}_{gstin}_{month_year[1]}_{month_year[0]}_{state_name}"
        if not _check_for_export_to_excel(page):
            log(f"Excel sheet not found for: {file_name}...")
            return
        else:
            log(f"✅ Excel sheet found for: {file_name}, attempting Excel download")

        # Only now expect a download
        with page.expect_download(timeout=DEFAULT_TIMEOUT) as download_info:
            page.click("#ctl00_ContentPlaceHolder1_btn_export_excel")
        # Save downloaded file
        download = download_info.value
        file_path = f"{downloads_dir}/{file_name}.xls"
        download.save_as(file_path)
        log(f"✅ Successfully downloaded data for {file_name}")
        # time.sleep(10)
    except Exception as e:
        log(f"❌ Exception in downloading Excel for {file_name}: {e}")


def download_EWB_for_gstin(page: Page, gstin: str, in_out_prefix: str, downloads_dir: str, month_year_tuple_list):
    """
    Download Excel reports for a specific GSTIN by iterating through all buyer states.
    Args:
        page: Playwright page object (EWB MIS page)
        gstin: GSTIN number to search for
        in_out_prefix: 'Out' or 'In' selection
        downloads_dir: Directory to save downloaded files
    """
    try:
        # Step 1: Select radio button (Outward/Inward) - do this once
        radio_selector = _get_radio_button_selector(in_out_prefix)
        page.wait_for_selector(radio_selector, timeout=DEFAULT_TIMEOUT)
        page.click(radio_selector)
        # page.wait_for_timeout(2000)
        log(f"Selected {in_out_prefix} radio button")

        # Step 3: Enter GSTIN - do this once
        gstin_selector = 'input[name="ctl00$ContentPlaceHolder1$txt_gstin"]'
        page.wait_for_selector(gstin_selector, timeout=DEFAULT_TIMEOUT)
        page.type(gstin_selector, gstin)
        # Small delay to ensure selection is registered
        # page.wait_for_timeout(2000)

        for month_year in month_year_tuple_list:
            last_day = get_days_in_month(month_year)
            # Step 2: Set Date From and To - do this once
            _set_date_fields_exact(page, last_day, month_year)
            # page.wait_for_timeout(2000)
            
            # Step 4: Iterate through each state option
            for state_value, state_name in state_options.items():
                if state_value == "0":  # Skip default "Select State" option
                    continue
                log(f"Checking state group: {state_value} : ({state_name})")
                try:
                    # Select the state
                    state_dropdown = 'select[name="ctl00$ContentPlaceHolder1$ddl_gstinstcode"]'
                    page.wait_for_selector(state_dropdown, timeout=DEFAULT_TIMEOUT)
                    page.select_option(state_dropdown, value=state_value)
                    # Small delay to ensure selection is registered
                    # page.wait_for_timeout(1000)
                    #TODO TIMEOUT
                    # Click GO button and check for data
                    _click_go_and_download_excel(page, gstin, state_name, month_year, downloads_dir, in_out_prefix)
                    # Small delay between states
                    # page.wait_for_timeout(2000)
                except Exception as state_error:
                    log(f"❌ Error processing state: {state_name}. :: {str(state_error)}")
                    continue
            log(f"Completed processing all states for GSTIN: {gstin} for {month_year[0]}_{month_year[1]}")
    except Exception as e:
        log(f"❌ Error processing GSTIN: {gstin} for {month_year[0]}_{month_year[1]}: {str(e)}")


def xls_to_xlsx(path, gst_id):
    """Converts .xls files to .xlsx in the specified path for a given GSTIN."""
    log("***Starting .xls to .xlsx conversion***")
    file_list1 = glob(os.path.join(path, f"In_{gst_id}*.xls"))
    file_list2 = glob(os.path.join(path, f"Out_{gst_id}*.xls"))
    file_list = file_list1 + file_list2

    if not file_list:
        log(f"No .xls files found for conversion in {path}.")
        return

    try:
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        excel.Visible = False # Run Excel in background

        for file in file_list:
            log(f"Converting file:{os.path.basename(file)}")
            wb = excel.Workbooks.Open(file)
            wb.SaveAs(file + "x", FileFormat=51) # FileFormat 51 is for .xlsx
            wb.Close()
        log("*** ✅ .xls to .xlsx Conversion was successful***")
    except Exception as e:
        log(f"❌ Error during .xls to .xlsx conversion: {e}")
    finally:
        if 'excel' in locals() and excel:
            excel.Quit() # Use .quit() to properly close Excel


def xlsx_merge(path, gst_id):
    """Merges all In_GSTIN_*.xlsx and Out_GSTIN_*.xlsx files into a single Merged_GSTIN.xlsx."""
    file_list1 = glob(os.path.join(path, f"In_{gst_id}*.xlsx"))
    file_list2 = glob(os.path.join(path, f"Out_{gst_id}*.xlsx"))
    file_list = file_list1 + file_list2

    if not file_list:
        log(f"No .xlsx files found for merging in {path} for GSTIN: {gst_id}.")
        return

    excl_list = []
    for file in file_list:
        log(f"Merging file: {os.path.basename(file)}")
        try:
            df = pd.read_excel(file)
            if not df.empty:
                excl_list.append(df)
        except Exception as e:
            log(f"❌ Error reading {os.path.basename(file)}: {e}")
            continue

    if not excl_list:
        log("No valid Excel files to merge.")
        return

    excl_merged = pd.concat(excl_list, ignore_index=True, sort=False)
    # excl_merged.drop_duplicates(inplace=True)
    l = len(excl_merged.axes[0])

    output_file = os.path.join(path, f'Merged_{gst_id}.xlsx')
    excl_merged.to_excel(output_file, index=False)
    log(f"✅ EWB In & Out files merge was successful and total number of rows are {l} for GSTIN: {gst_id}.")
    del excl_merged
    del excl_list
    del df
    gc.collect()


def ewbextract_stock_stmt(page, ewbs, dpath):
    """
    Alternative version with more reliable dialog handling.
    """
    total = len(ewbs)
    log(f"Starting EWB details extraction for {total} EWBs...")

    for idx, ewb_no in enumerate(ewbs, start=1):
        try:
            log(f"Starting EWB details extraction for EWB: {ewb_no}")
            url = f"https://mis.ewaybillgst.gov.in/Verify/EwayBillPrint.aspx?ewb_no={ewb_no}&cal=1"
            page.goto(url, wait_until="networkidle", timeout=DEFAULT_TIMEOUT)

            # Wait for all required elements and extract text
            page.locator('#ctl00_ContentPlaceHolder1_lblApxDistDetails').wait_for(timeout=DEFAULT_TIMEOUT)
            dist = page.locator('#ctl00_ContentPlaceHolder1_lblApxDistDetails').text_content()
            
            page.locator('#ctl00_ContentPlaceHolder1_lblTransType').wait_for(timeout=DEFAULT_TIMEOUT)
            trans = page.locator('#ctl00_ContentPlaceHolder1_lblTransType').text_content()
            
            page.locator('#ctl00_ContentPlaceHolder1_txtGenBy').wait_for(timeout=DEFAULT_TIMEOUT)
            frm_addr = page.locator('#ctl00_ContentPlaceHolder1_txtGenBy').text_content()
            
            page.locator('#ctl00_ContentPlaceHolder1_txtSypplyTo').wait_for(timeout=DEFAULT_TIMEOUT)
            to_addr = page.locator('#ctl00_ContentPlaceHolder1_txtSypplyTo').text_content()
            # Check for main item list
            if page.is_visible('#ctl00_ContentPlaceHolder1_GVItemList', timeout=DEFAULT_TIMEOUT):
                html = page.locator('#ctl00_ContentPlaceHolder1_GVItemList').evaluate("el => el.outerHTML")
                df = pd.read_html(StringIO(html))[0]
                df = df.assign(ewb=ewb_no, Dist=dist, Trans=trans, From=frm_addr, To=to_addr)
                df.to_excel(os.path.join(dpath, f"{ewb_no}.xlsx"), index=False)
                log(f"[{idx}/{total}] Downloaded item list for EWB: {ewb_no}")

            else:
                log(f"[{idx}/{total}] Main table not found for EWB: {ewb_no}, checking IRN fallback...")
                if page.is_visible('#ctl00_ContentPlaceHolder1_btn_irn', timeout=DEFAULT_TIMEOUT):
                    try:
                        # Use expect_event to handle dialog
                        with page.expect_event('dialog', timeout=DEFAULT_TIMEOUT) as dialog_info:
                            page.locator('#ctl00_ContentPlaceHolder1_btn_irn').click()
                        
                        # Dialog appeared
                        dialog = dialog_info.value
                        dialog.accept()
                        log(f"[{idx}/{total}] IRN dialog handled for EWB: {ewb_no}")
                        
                        # Create dummy data
                        df = pd.DataFrame([{
                            'ewb': ewb_no, 'Dist': dist, 'Trans': trans,
                            'From': frm_addr, 'To': to_addr,
                            'HSN Code': '', 'Quantity': ''
                        }])
                        df.to_excel(os.path.join(dpath, f"{ewb_no}_dist.xlsx"), index=False)
                        log(f"[{idx}/{total}] Created dummy data for EWB: {ewb_no}")
                        
                    except Exception:
                        # No dialog appeared, check for IRN table
                        if page.is_visible('#ctl00_ContentPlaceHolder1_grd_items', timeout=DEFAULT_TIMEOUT):
                            html = page.locator('#ctl00_ContentPlaceHolder1_grd_items').evaluate("el => el.outerHTML")
                            df = pd.read_html(StringIO(html))[0]
                            df = df.assign(ewb=ewb_no, Dist=dist, Trans=trans, From=frm_addr, To=to_addr)
                            df.to_excel(os.path.join(dpath, f"{ewb_no}_irn.xlsx"), index=False)
                            log(f"[{idx}/{total}] Downloaded IRN item list for EWB: {ewb_no}")
                        else:
                            log(f"[{idx}/{total}] No item list found after IRN click for EWB: {ewb_no}")
                else:
                    log(f"[{idx}/{total}] No IRN button found for EWB: {ewb_no}")

        except Exception as e:
            log(f"[{idx}/{total}] ❌ Error processing EWB: {ewb_no}: {e}")


def xlsx_mergejoinsort_stock_stmt(dpath, mfile, edfm_main):
    """
    Merges, joins, sorts EWB data and prepares the stock statement.
    Args:
        dpath (str): The GSTIN-specific download directory.
        mfile (str): Merged file prefix (e.g., 'Merged_GSTIN').
        edfm_main (pd.DataFrame): The main merged EWB DataFrame (from Merged_GSTIN.xlsx).
    """
    try:
        file_list = glob(os.path.join(dpath, "[0-9]"*12 + ".xlsx")) # EWB files
        file_list2 = glob(os.path.join(dpath, "[0-9]"*12 + "_irn.xlsx")) # IRN files
        file_list3 = glob(os.path.join(dpath, "[0-9]"*12 + "_dist.xlsx")) # Dist files (dummy for IRN alerts)
        
        excl_list = []
        excl_list2 = []
        excl_list3 = []

        # Process EWB files
        if len(file_list) > 0:
            for file in file_list:
                log(f"Merging EWB file: {os.path.basename(file)}")
                try:
                    excl_list.append(pd.read_excel(file))
                    os.remove(file) # Clean up
                except Exception as e:
                    log(f"❌ Error reading/removing EWB file {os.path.basename(file)}: {e}")
            excl_merged = pd.concat(excl_list, ignore_index=True) if excl_list else pd.DataFrame()
            excl_merged = excl_merged[['HSN Code', 'Quantity', 'Taxable Amount Rs.', 'Dist', 'Trans', 'From', 'To', 'ewb']]
            excl_merged = excl_merged.rename(columns={'Taxable Amount Rs.': 'Taxable_Amt'})
            log(f"EWB Merge successful. Rows: {len(excl_merged)}")
            excl_list.clear()
            del excl_list
        else:
            excl_merged = pd.DataFrame()

        # Process IRN files
        if len(file_list2) > 0:
            for file in file_list2:
                log(f"Merging IRN file: {os.path.basename(file)}")
                try:
                    excl_list2.append(pd.read_excel(file))
                    os.remove(file) # Clean up
                except Exception as e:
                    log(f"❌ Error reading/removing IRN file {os.path.basename(file)}: {e}")
            excl_merged2 = pd.concat(excl_list2, ignore_index=True) if excl_list2 else pd.DataFrame()
            excl_merged2['Quantity'] = excl_merged2['Quantity'].astype(str) + ' ' + excl_merged2['Unit']
            excl_merged2 = excl_merged2[['HSN Code', 'Quantity', 'Taxable Amount(Rs)', 'Dist', 'Trans', 'ewb']]
            excl_merged2 = excl_merged2.rename(columns={'Taxable Amount(Rs)': 'Taxable_Amt'})
            log(f"EWB IRN Merge successful. Rows: {len(excl_merged2)}")
            excl_list2.clear()
            del excl_list2
        else:
            excl_merged2 = pd.DataFrame()
            
        # Process Dist files
        if len(file_list3) > 0:
            for file in file_list3:
                log(f"Merging Dist file: {os.path.basename(file)}")
                try:
                    excl_list3.append(pd.read_excel(file))
                    os.remove(file) # Clean up
                except Exception as e:
                    log(f"❌ Error reading/removing Dist file {os.path.basename(file)}: {e}")
            excl_merged3 = pd.concat(excl_list3, ignore_index=True) if excl_list3 else pd.DataFrame()
            excl_merged3['Quantity'] = excl_merged3['Quantity'].astype(str)
            excl_merged3 = excl_merged3[['HSN Code', 'Quantity', 'Dist', 'Trans', 'From', 'To', 'ewb']]
            log(f"✅ EWB IRN Dist Merge successful. Rows: {len(excl_merged3)}")
            excl_list3.clear()
            del excl_list3
        else:
            excl_merged3 = pd.DataFrame()

        excl_ewbs = pd.concat([excl_merged, excl_merged2, excl_merged3], ignore_index=True)
        del excl_merged, excl_merged2, excl_merged3
        gc.collect()
        log(f"All EWB HSN combinations: {len(excl_ewbs)}")

        edf = excl_ewbs.copy()
        edf['Quantity'] = edf['Quantity'].astype(str)
        edf['Quantity'] = edf['Quantity'].str.upper() + ' '
        edf['Qty'] = edf['Quantity'].str.split(' ', expand=True)[0]
        edf['Qty'] = edf['Qty'].astype(float)
        edf['Unit'] = edf['Quantity'].str.split(' ', expand=True)[1]
        edf['Unit'] = edf['Unit'].str.upper()
        edf['HSNCode'] = edf['HSN Code'].astype(str)
        edf = edf.drop(['Quantity', 'HSN Code'], axis=1)
        edf = edf.sort_values(by=['ewb', 'HSNCode']).reset_index()
        final = pd.merge(edfm_main, edf, on='ewb', how='inner')
        final['HSN4'] = final['HSN Code'].astype(str).str[:4]
        final['HSNCode'] = pd.to_numeric(final['HSN Code'], errors='coerce').fillna(0).astype(int)
        final = final.drop(['index'], axis=1)
        final[['EWB No. & Dt.','DateTime']] = final['EWB No. & Dt.'].str.split('-',expand=True)
        final.drop(['EWB No. & Dt.'], axis=1)
        final['DateTime'] = final['DateTime'].str.strip()
        final['DateTime']= pd.to_datetime(final['DateTime'], format='%d/%m/%Y %H:%M:%S')
        final.sort_values(by='DateTime', inplace=True)
        final.loc[final['Qty'] > 100, 'Qty'] = final['Qty']/1000


        # Apply quantity correction (if Qty > 100, divide by 1000)
        final.loc[final['Qty'] > 100, 'Qty'] = final['Qty']/1000

        # Define Purchase/Sale columns based on GSTIN
        final['Purchase from'] = final['From GSTIN & Name']
        final['Sale To'] = final['To GSTIN & Name']
        final['Pur_Value'] = final['Sale_Value'] = final['Assess Val.']
        final['Pur_TaxVal'] = final['Sale_TaxVal'] = final['Tax Val.']
        final['Pur_Vehicle'] = final['Sale_Vehicle'] = final['Latest Vehicle No.']
        final['Pur_Qty'] = final['Sale_Qty'] = final['Qty']

        
        # Adjust values based on whether it's a purchase or sale for the current GSTIN
        final.loc[final['Purchase from'].astype(str).str.contains(mfile.replace('Merged_','')), ['Pur_Value', 'Pur_TaxVal', 'Pur_Vehicle', 'Pur_Qty']] = [0, 0, '', 0]
        final.loc[final['Sale To'].astype(str).str.contains(mfile.replace('Merged_','')), ['Sale_Value', 'Sale_TaxVal', 'Sale_Vehicle', 'Sale_Qty']] = [0, 0, '', 0]

        final['Pur_Qty'] = final['Pur_Qty'].fillna(0)
        final['Sale_Qty'] = final['Sale_Qty'].fillna(0)

        final['0B'] = 0
        final['Total Stock'] = 0
        final['CB'] = 0
        final[' '] = ''

        final = final.drop(['From GSTIN & Name','To GSTIN & Name','EWB No. & Dt.','Doc No. & Dt.','Assess Val.','Tax Val.','HSNCode','HSN Desc.','Latest Vehicle No.','Qty','From Place & Pin','To Place & Pin'], axis=1)

        distinct_hsns = final['HSN4'].unique()

        excel_file_path = dpath + '/' + mfile + '_stockstmnt.xlsx'
        with pd.ExcelWriter(excel_file_path) as writer:
            for value in distinct_hsns:
                try:
                    final_hsn = final[final['HSN4'] == value].copy()
                    final_hsn.loc[:, 'S.No'] = np.arange(1, len(final_hsn) + 1)
                    final_hsn['CB'] = (final_hsn['Pur_Qty'] - final_hsn['Sale_Qty']).cumsum()
                    final_hsn['0B'] = final_hsn['CB'].shift(fill_value=0)
                    final_hsn['Total Stock'] = final_hsn['0B'] + final_hsn['Pur_Qty']
                    final_hsn['0B'] = final_hsn['0B'].round(2)
                    final_hsn['CB'] = final_hsn['CB'].round(2)
                    final_hsn['Total Stock'] = final_hsn['Total Stock'].round(2)
                    start_excel_row = 2
                    final_hsn['EWB Toll'] = [f'=HYPERLINK("#"&CELL("address",INDEX(TollData!A:A,MATCH(D{excel_row},TollData!A:A,0))),D{excel_row})' for excel_row in range(start_excel_row, start_excel_row + len(final_hsn))]
                    final_hsn['States in which vehicle movement exists'] = [f'=VLOOKUP(D{excel_row},TollUniq!A:B,{excel_row},FALSE)' for excel_row in range(start_excel_row, start_excel_row + len(final_hsn))]
                    final_hsn = final_hsn[['S.No','DateTime','0B','EWB No.','EWB Toll','HSN Code','Trans', 'Purchase from', 'From', 'Pur_Qty','Pur_Value','Pur_TaxVal','Pur_Vehicle','Total Stock',' ','Sale To', 'To', 'Sale_Qty','Sale_Value','Sale_TaxVal','Sale_Vehicle','CB','Dist','States in which vehicle movement exists']]
                    #final_hsn = final_hsn[['S.No','DateTime','0B','EWB No.','HSN Code', 'Purchase from', 'Pur_Qty','Pur_Value','Pur_TaxVal','Pur_Vehicle','Total Stock',' ','Sale To', 'Sale_Qty','Sale_Value','Sale_TaxVal','Sale_Vehicle','CB','Dist']]
                    final_hsn.to_excel(writer, sheet_name=value, index=False)
                    print(f"*** Stock statement creation for HSN: {value} is complete ***")
                except Exception as e:
                    log(f"❌ Error creating stock statement for for HSN: {value} : {e}")
            del final, final_hsn, edf
            gc.collect()
            log(f"*** ✅ Stock statement creation for {mfile.replace('Merged_','')} is complete ***")
    except Exception as e:
        log(f"❌ Error creating stock statement Excel file for {mfile}: {e}")


def xlsxsheetmerge(mgstin, dpath):
    """
    Merges all sheets from the stock statement file for the given GSTIN.
    Args:
        mgstin (str): GSTIN ID.
        dpath (str): The GSTIN-specific download directory.
    """
    try:
        path_obj = Path(dpath)
        file_list = list(path_obj.glob(f"Merged_{mgstin}_stockstmnt.xlsx"))
        log(f"Files found for sheet merge: {file_list}")

        excl_list = []

        for file in file_list:
            log(f"Processing sheets from {file.name}")
            try:
                df_dict = pd.read_excel(file, sheet_name=None)
                for sheet, sheet_df in df_dict.items():
                    sheet_df['SheetName'] = sheet  # Add sheet name column
                    excl_list.append(sheet_df)
                del df_dict
                gc.collect()
            except Exception as e:
                log(f"❌ Error processing {file.name}: {e}")

        if not excl_list:
            log("No sheets processed for merging.")
            return

        excl_merged = pd.concat(excl_list, ignore_index=True)
        total_rows = len(excl_merged)

        new_excl_merged = excl_merged.drop_duplicates()
        deduped_rows = len(new_excl_merged)

        output_file = path_obj / f'Merged_{mgstin}_stockstmntall.xlsx'
        new_excl_merged.to_excel(output_file, index=False)

        log(f"✅ Sheet merge successful. Rows before/after duplicates: {total_rows} and {deduped_rows}")

        # Final cleanup
        del excl_list, excl_merged, new_excl_merged
        gc.collect()
    except Exception as e:
        log(f"❌ Error while function call xlsxsheetmerge() for GSTIN: {mgstin}")


def xlsx_mergejoinsort_toll_details(dpath, mfile):
    """
    Merges, sorts toll data and appends to the stock statement file.
    Args:
        dpath (str): The GSTIN-specific download directory.
        mfile (str): Merged file prefix (e.g., 'Merged_GSTIN').
    """
    try:
        file_list = glob(os.path.join(dpath, "[0-9]"*12 + "_toll.xlsx"))
        
        excl_list = []
        if len(file_list) > 0:
            for file in file_list:
                log(f"Merging toll file: {os.path.basename(file)}")
                try:
                    excl_list.append(pd.read_excel(file))
                    os.remove(file) # Clean up
                except Exception as e:
                    log(f"❌ Error reading/removing toll file: {os.path.basename(file)}: {e}")
            excl_merged = pd.concat(excl_list, ignore_index=True) if excl_list else pd.DataFrame()
            log(f"✅ EWB toll merge successful. Rows: {len(excl_merged)}")
        else:
            excl_merged = pd.DataFrame()
            log("No toll files found to merge.")
            return # Exit if no toll data was found

        # Reorder columns to have 'ewb' first
        if 'ewb' in excl_merged.columns:
            cols = excl_merged.columns.tolist()
            cols.insert(0, cols.pop(cols.index('ewb'))) # Move 'ewb' to front
            excl_ewbs = excl_merged[cols]
        else:
            excl_ewbs = excl_merged # No 'ewb' column, proceed as is
        log(f"Length of All EWB toll combinations: {len(excl_ewbs)}")

        existing_excel_path = os.path.join(dpath, f'{mfile}_stockstmnt.xlsx')
        
        if not os.path.exists(existing_excel_path):
            log(f"❌ Error: Stock statement file not found at {existing_excel_path}. Cannot append Toll Data.")
            return

        try:
            with pd.ExcelWriter(existing_excel_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                excl_ewbs.to_excel(writer, index=False, sheet_name='TollData')

                # Prepare TollUniq sheet
                if 'State' in excl_merged.columns and 'ewb' in excl_merged.columns:
                    excl_merged_unique_states = excl_merged[['ewb', 'State']].drop_duplicates()
                    excl_merged_unique_states = excl_merged_unique_states[excl_merged_unique_states['State'].notna() & (excl_merged_unique_states['State'] != '')]
                    result_unique_states = excl_merged_unique_states.groupby('ewb')['State'].agg(','.join).reset_index()
                    result_unique_states.to_excel(writer, index=False, sheet_name='TollUniq')
                    log("TollData and TollUniq sheets appended to existing Excel file!")
                else:
                    log("Skipping TollUniq sheet creation: 'State' or 'ewb' column missing in toll data.")
        except Exception as e:
            log(f"❌ Error appending toll sheets to Excel file for file: {mfile}: {e}")
    except Exception as e:
        log(f"❌ Error whiile function call xlsx_mergejoinsort_toll_details() for file: {mfile}: {e}")


def ewb_extract_toll_details(page, ewbs: list, dpath: str):
    """
    Args:
        page: The Playwright sync Page instance.
        ewbs (list): List of EWB numbers to extract toll data for.
        dpath (str): The GSTIN-specific download directory.
        timeout (int): Timeout in milliseconds (default: 30000ms = 30s).
    """
    lewb = len(ewbs)
    log(f"Starting EWB toll extraction for {lewb} EWBs...")
    
    for i, ewb in enumerate(ewbs):
        try:
            page.goto(
                f"https://mis.ewaybillgst.gov.in/RFID_Reports/Ewb_rpt.aspx?id=1&ewayno={ewb}",
                timeout=_5_MIN_TIMEOUT,
                wait_until='load'
            )
            table_selector = "#ctl00_ContentPlaceHolder1_grd_tolldtls"
            try:
                page.wait_for_selector(table_selector, timeout=DEFAULT_TIMEOUT)
                html = page.locator(table_selector).evaluate("el => el.outerHTML")
                if html:
                    # Parse HTML table into DataFrame
                    df = pd.read_html(StringIO(html))[0]
                    df['ewb'] = ewb
                    cols = df.shape[1]
                    
                    if cols <= 2:  # Typically means no detailed toll data
                        log(f"Toll data not found (or incomplete) for ewb {ewb} -> Progress {i+1}/{lewb}")
                    else:
                        dfile = os.path.join(dpath, f"{ewb}_toll.xlsx")
                        df.to_excel(dfile, index=False)
                        log(f"Downloaded ewb toll table for ewb {ewb} -> Progress {i+1}/{lewb}")
                else:
                    log(f"Could not extract HTML toll data for ewb {ewb} -> Progress {i+1}/{lewb}")
                    
            except TimeoutError:
                log(f"Toll data table not found for ewb {ewb} (timeout) -> Progress {i+1}/{lewb}")
            
        except TimeoutError as e:
            log(f"❌ Timeout error for EWB {ewb}: {e}")
        except Exception as e:
            log(f"❌ Error extracting toll data for EWB {ewb}: {e}")


def main():
    # Load config file
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        config = json.load(f)
    username = config["username"]
    password = config["password"]
    gstins = config["gstins"]
    # from_date = config["start_date"]
    # to_date = config["end_date"]
    start_month = config["start_month"]
    end_month = config["end_month"]
    start_year = config["start_year"]
    end_year = config["end_year"]
    extract_ewb_data_flag = config["extract_ewb_data_flag"]
    prepare_stock_statement_flag = config["prepare_stock_statement_flag"]
    check_toll_data_flag = config["check_toll_data_flag"]
    
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=False, args=["--start-maximized"])
            context = browser.new_context(accept_downloads=True)
            # ewb_page = context.new_page()
            page = context.new_page()
            #Login and navigate to EWB MIS portal
            try:
                ewb_page = login_and_open_ewb_mis(page, context, username, password)
            except Exception as e:
                log(f"Login or EWB MIS navigation failed: {e}")
                context.close()
                return
            month_year_tuple_list = get_month_year_range(start_month, start_year, end_month, end_year)
            log(month_year_tuple_list)
            
            # Loop over GSTINs and download E-Way Bill
            if extract_ewb_data_flag:
                for gstin in gstins:
                    try:
                        log(f"Starting to extract EWB for GSTIN: {gstin}")
                        downloads_dir = os.path.abspath(f"./output/{gstin}")
                        os.makedirs(downloads_dir, exist_ok=True)
                        download_EWB_for_gstin(ewb_page, gstin, _IN_, downloads_dir, month_year_tuple_list)
                        download_EWB_for_gstin(ewb_page, gstin, _OUT_, downloads_dir, month_year_tuple_list)
                        xls_to_xlsx(downloads_dir, gstin)
                        xlsx_merge(downloads_dir, gstin)
                        log(f"✅ E-Way Bill extraction and merge complete for GSTIN: {gstin}.")
                    except Exception as e:
                        log(f"❌ Error while E-Way Bill extraction and merge for {gstin}: {e}")
            else: 
                log(f"Skipping downloading E-Way bills from GST portal as extract_ewb_data_flag is False.")

            # Loop over GSTINs and prepare stock statement
            if prepare_stock_statement_flag:
                for gstin in gstins:
                    try:
                        log(f"Preparing Stock Statement for GSTIN: {gstin}")
                        downloads_dir = os.path.abspath(f"./output/{gstin}")
                        os.makedirs(downloads_dir, exist_ok=True)
                        mfile = 'Merged_' + gstin
                        merged_ewb_path = os.path.join(downloads_dir, mfile + '.xlsx')
                        
                        if not os.path.exists(merged_ewb_path):
                            log(f"❌ Error: Merged EWB file not found for {gstin} at {merged_ewb_path}. Skipping Stock Statement.")
                        else:
                            edfm = pd.read_excel(merged_ewb_path)
                            edfm['ewb'] = edfm['EWB No.']
                            ewbs = edfm['ewb'].tolist()
                            
                            ewbextract_stock_stmt(ewb_page, ewbs, downloads_dir)
                            xlsx_mergejoinsort_stock_stmt(downloads_dir, mfile, edfm)
                            xlsxsheetmerge(gstin, downloads_dir)
                            log(f"✅ Stock Statement preparation complete for GSTIN: {gstin}.")
                    except Exception as e:
                        log(f"❌ Error while stock statement preparation for {gstin}: {e}")
            else: 
                log(f"Skipping preparing stock statement from GST portal as prepare_stock_statement_flag is False.")

            if check_toll_data_flag:
                for gstin in gstins:
                    try:
                        log(f"Checking Toll data for GSTIN: {gstin}...")
                        downloads_dir = os.path.abspath(f"./output/{gstin}")
                        os.makedirs(downloads_dir, exist_ok=True)
                        mfile = 'Merged_' + gstin
                        merged_ewb_path = os.path.join(downloads_dir, mfile + '.xlsx')
                    
                        if not os.path.exists(merged_ewb_path):
                            log(f"❌ Error: Merged EWB file not found for {gstin} at {merged_ewb_path}. Skipping Toll Check.")
                        else:
                            edfm = pd.read_excel(merged_ewb_path)
                            edfm['ewbno'] = edfm['EWB No.']
                            ewbs = edfm['ewbno'].tolist()
                            
                            ewb_extract_toll_details(ewb_page, ewbs, downloads_dir)
                            xlsx_mergejoinsort_toll_details(downloads_dir, mfile)
                            log(f"✅ Toll details creation complete for {gstin}.")
                    except Exception as e:
                        log(f"❌ Error while Toll details creation for {gstin}: {e}")
            else: 
                log(f"Skipping toll data from GST portal as check_toll_data_flag is False.")

            log("~*~ ✅All GSTINs processed successfully✅ ~*~")
            # ewb_page.goto("https://ewaybillgst.gov.in/mainmenu.aspx")
            time.sleep(600000)
            context.close()
    except Exception as e:
        log(f"❌ Fatal error during browser automation: {e}")
        

if __name__ == "__main__":
    main()

