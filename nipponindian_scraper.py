import logging
import requests
import pandas as pd
from datetime import datetime
import time as tm
import xlwings as xw


def log_to_sheet(message):
    # Open the workbook and sheet
    wb = xw.Book.caller()
    sheet = wb.sheets['Sheet1']

    # Find the next available row starting from row 11
    last_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
    if last_row < 16:
        last_row = 16  # Ensure logs always start at row 11

    # Write the log message in the first column below the last log
    sheet.range(f'A{last_row + 1}').value = message


def run_scraper(start_time, end_time, crawl_gap):
    # Initialize logging
    log_to_sheet("Script started")

    cookies = {
    # your cookies
}

    headers = {
   # your address
}
    try:
        current_time = datetime.now().time()
        start_time_obj = datetime.strptime(start_time, '%H:%M').time()
        end_time_obj = datetime.strptime(end_time, '%H:%M').time()

        log_to_sheet("Start time & End time fetched successfully")
        while current_time < end_time_obj:
            if current_time >= start_time_obj:
                response = requests.post('https://investeasy.nipponindiaim.com/Online/Realtime/DetailsFill',
                                         cookies=cookies, headers=headers)
                page = response.json()

                extracted_data = [
                    {
                        'Date': datetime.now().date(),
                        'Time': datetime.now().strftime('%H:%M:%S'),
                        'SchName': item['SchName'],
                        'CNav': item['CNav'],
                        'PNav': item['PNav'],
                        'NCvalue': item['NCvalue'],
                        'PChange': item['PChange'],
                        'Link': item['Link'],
                        'Realdt': item['Realdt'],
                        'Category': item['Category']
                    }
                    for item in page['RVDetailsList']
                ]

                df = pd.DataFrame(extracted_data)
                wb = xw.Book.caller()
                sheet = wb.sheets['CrawlData']

                # Find the next empty row in CrawlData sheet
                last_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row + 1
                sheet.range(f'A{last_row}').value = df.values.tolist()

                wb.save()
                log_to_sheet("Data Fetched")

            # Wait for the specified crawl gap (in minutes)
            tm.sleep(int(crawl_gap) * 60)
            current_time = datetime.now().time()

        log_to_sheet("Crawl Completed")
    except Exception as e:
        log_to_sheet(f"An error occurred: {str(e)}")
