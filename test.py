import logging
from bs4 import BeautifulSoup
import requests
import pandas as pd
from datetime import datetime
import time as tm
import os
import xlwings as xw
import openpyxl

# Set up logging
logging.basicConfig(
    filename='nipponindia_scraper.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

# Function to append data to the existing Excel file without overwriting
def append_to_excel(df, filename, sheet_name='Data'):
    try:
        # Check if the file already exists
        if not os.path.isfile(filename):
            # If it doesn't exist, create a new file with headers
            df.to_excel(filename, index=False, sheet_name=sheet_name)
        else:
            # If the file exists, append to the existing file
            with pd.ExcelWriter(filename, engine='openpyxl', mode='a') as writer:
                if sheet_name in writer.sheets:
                    # Append data to the existing sheet
                    startrow = writer.sheets[sheet_name].max_row
                else:
                    startrow = 0
                df.to_excel(writer, index=False, sheet_name=sheet_name, startrow=startrow, header=startrow == 0)
    except Exception as e:
        logging.error(f"Failed to append to Excel: {str(e)}")


# Function to run the scraping
def run_scraper(start_time_str, end_time_str):
    try:
        # Convert start and end times to time objects
        start_time = datetime.strptime(start_time_str, '%H:%M:%S').time()
        end_time = datetime.strptime(end_time_str, '%H:%M:%S').time()

        cookies = {
            # Add your cookie details here
        }

        headers = {
            # Add your header details here
        }

        # Loop to run the code every minute within the time frame
        while True:
            current_time = datetime.now().time()

            if start_time <= current_time <= end_time:
                try:
                    # Sending the request
                    response = requests.post('https://investeasy.nipponindiaim.com/Online/Realtime/DetailsFill',
                                             cookies=cookies, headers=headers)
                    page = response.json()

                    # Extracting the data
                    extracted_data = [
                        {
                            'Date': datetime.now().date(),  # Adding a date column
                            'Time': datetime.now().strftime('%H:%M:%S'),  # Adding a time column
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

                    # Converting the data to a DataFrame
                    df = pd.DataFrame(extracted_data)

                    # Appending the data to the Excel file
                    filename = "nippon_execution.xlsm"
                    append_to_excel(df, filename)

                    # Log the successful data save
                    logging.info(f"Data successfully appended to {filename}")

                except Exception as e:
                    # Log any errors that occur
                    logging.error(f"An error occurred while scraping: {str(e)}")

            # Wait for 1 minute before the next iteration
            tm.sleep(60)

    except Exception as e:
        logging.error(f"An error occurred while running the scraper: {str(e)}")


# Function to be called from Excel
def excel_scraper_control():    
    wb = xw.Book.caller()  # Connect to the calling Excel workbook
    sheet = wb.sheets['Sheet1']  # Adjust to your sheet name

    start_time_str = sheet.range('L2').value  # Assuming start time is in cell L2
    end_time_str = sheet.range('L3').value  # Assuming end time is in cell L3

    sheet.range('L4').value = "Script Running..."  # Update status
    run_scraper(start_time_str, end_time_str)
    sheet.range('L4').value = "Script Completed"  # Update status after completion


if __name__ == "__main__":
    xw.Book("nippon_execution.xlsm").set_mock_caller()  # Replace with your Excel filename
    excel_scraper_control()
    input("Press Enter to exit...")