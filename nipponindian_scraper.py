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
    'at_check': 'true',
    'AMCVS_C68C337B54EA1B460A4C98A1%40AdobeOrg': '1',
    'AMCV_C68C337B54EA1B460A4C98A1%40AdobeOrg': '179643557%7CMCIDTS%7C19991%7CMCMID%7C88569985073849471860196101119611606174%7CMCAAMLH-1727769982%7C12%7CMCAAMB-1727769982%7CRKhpRz8krg2tLO6pguXWp5olkAcUniQYPHaMWWgdJ3xzPWQmdj0y%7CMCOPTOUT-1727172382s%7CNONE%7CvVersion%7C5.5.0',
    'mbox': 'session#578d7f4dc84442e49abb777f5963f3d5#1727167043|PC#578d7f4dc84442e49abb777f5963f3d5.41_0#1790409983',
    '_ga_9LDNS8Y4ZW': 'GS1.1.1727165182.1.0.1727165182.0.0.0',
    '_ga_Z5N4HF2573': 'GS1.1.1727165182.1.0.1727165182.0.0.0',
    'ASP.NET_SessionId': 'vse5ankw0uh5fzo5ezbcd0gm',
    'NIMF': 'rd7o00000000000000000000ffff0a290761o80',
    'TS01f4aefd': '0176bf02ace97dc36168ba7c68ab5328dd4e08787d927cda6eb0849878534876ead9a06fa089bf3b303a2df3f0468955ec4a4eb448438049ae2194ccd8b5a222129759adfbb42196e4f9b10464e70bbe0a60b9bd14',
    'gpv': 'mf%3Afundsandperformance%3Apages%3Ainav',
    's_cc': 'true',
    '_hjSessionUser_5078605': 'eyJpZCI6IjZhMDM4MzU5LTMwMGYtNTdhMi1iNzZmLTQ5YzBlN2E4OWE0YiIsImNyZWF0ZWQiOjE3MjcxNjUxODk4MjMsImV4aXN0aW5nIjpmYWxzZX0=',
    '_hjSession_5078605': 'eyJpZCI6IjhiZGU2MzBjLWQ1ZTUtNDkwMi1iOTczLThlYjRiZDMwODY0MiIsImMiOjE3MjcxNjUxODk4MjcsInMiOjAsInIiOjAsInNiIjowLCJzciI6MCwic2UiOjAsImZzIjoxLCJzcCI6MX0=',
    's_nr': '1727165202441-New',
    's_ppvl': 'mf%253Afundsandperformance%253Apages%253Ainav%2C31%2C31%2C746%2C1528%2C746%2C1536%2C864%2C1.25%2CP',
    '_fbp': 'fb.1.1727165212211.370110310976028504',
    '_gcl_au': '1.1.1855227823.1727165212',
    '_uetsid': 'f52aed607a4b11ef949449108f8fd736',
    '_uetvid': 'f52b3b407a4b11ef8e2823d14de77fe0',
    '_gid': 'GA1.2.1712692576.1727165213',
    '_gat_gtag_UA_9474483_24': '1',
    '_ga_NNCDXFQMC2': 'GS1.1.1727165212.1.0.1727165212.60.0.0',
    '_ga': 'GA1.1.965630197.1727165183',
    'TSe2513c34027': '08d45de36dab2000da0571195ce5141e05af077bf1dd1c4b08afc4deb4fb65eb3f4275d6c3c2139408dc33b5f3113000e343bcee4b1817fd0803a169e68eeb52559244dccad298b92cbbaacc8b122d4cb875095ab6c2953b1cc00c3cc2c168b5',
    's_ppv': 'mf%253Afundsandperformance%253Apages%253Ainav%2C31%2C31%2C750%2C794%2C746%2C1536%2C864%2C1.25%2CP',
}

    headers = {
    'Accept': 'application/json, text/javascript, */*; q=0.01',
    'Accept-Language': 'en-US,en;q=0.9',
    'Connection': 'keep-alive',
    # 'Content-Length': '0',
    'Content-Type': 'application/json',
    # 'Cookie': 'at_check=true; AMCVS_C68C337B54EA1B460A4C98A1%40AdobeOrg=1; AMCV_C68C337B54EA1B460A4C98A1%40AdobeOrg=179643557%7CMCIDTS%7C19991%7CMCMID%7C88569985073849471860196101119611606174%7CMCAAMLH-1727769982%7C12%7CMCAAMB-1727769982%7CRKhpRz8krg2tLO6pguXWp5olkAcUniQYPHaMWWgdJ3xzPWQmdj0y%7CMCOPTOUT-1727172382s%7CNONE%7CvVersion%7C5.5.0; mbox=session#578d7f4dc84442e49abb777f5963f3d5#1727167043|PC#578d7f4dc84442e49abb777f5963f3d5.41_0#1790409983; _ga_9LDNS8Y4ZW=GS1.1.1727165182.1.0.1727165182.0.0.0; _ga_Z5N4HF2573=GS1.1.1727165182.1.0.1727165182.0.0.0; ASP.NET_SessionId=vse5ankw0uh5fzo5ezbcd0gm; NIMF=rd7o00000000000000000000ffff0a290761o80; TS01f4aefd=0176bf02ace97dc36168ba7c68ab5328dd4e08787d927cda6eb0849878534876ead9a06fa089bf3b303a2df3f0468955ec4a4eb448438049ae2194ccd8b5a222129759adfbb42196e4f9b10464e70bbe0a60b9bd14; gpv=mf%3Afundsandperformance%3Apages%3Ainav; s_cc=true; _hjSessionUser_5078605=eyJpZCI6IjZhMDM4MzU5LTMwMGYtNTdhMi1iNzZmLTQ5YzBlN2E4OWE0YiIsImNyZWF0ZWQiOjE3MjcxNjUxODk4MjMsImV4aXN0aW5nIjpmYWxzZX0=; _hjSession_5078605=eyJpZCI6IjhiZGU2MzBjLWQ1ZTUtNDkwMi1iOTczLThlYjRiZDMwODY0MiIsImMiOjE3MjcxNjUxODk4MjcsInMiOjAsInIiOjAsInNiIjowLCJzciI6MCwic2UiOjAsImZzIjoxLCJzcCI6MX0=; s_nr=1727165202441-New; s_ppvl=mf%253Afundsandperformance%253Apages%253Ainav%2C31%2C31%2C746%2C1528%2C746%2C1536%2C864%2C1.25%2CP; _fbp=fb.1.1727165212211.370110310976028504; _gcl_au=1.1.1855227823.1727165212; _uetsid=f52aed607a4b11ef949449108f8fd736; _uetvid=f52b3b407a4b11ef8e2823d14de77fe0; _gid=GA1.2.1712692576.1727165213; _gat_gtag_UA_9474483_24=1; _ga_NNCDXFQMC2=GS1.1.1727165212.1.0.1727165212.60.0.0; _ga=GA1.1.965630197.1727165183; TSe2513c34027=08d45de36dab2000da0571195ce5141e05af077bf1dd1c4b08afc4deb4fb65eb3f4275d6c3c2139408dc33b5f3113000e343bcee4b1817fd0803a169e68eeb52559244dccad298b92cbbaacc8b122d4cb875095ab6c2953b1cc00c3cc2c168b5; s_ppv=mf%253Afundsandperformance%253Apages%253Ainav%2C31%2C31%2C750%2C794%2C746%2C1536%2C864%2C1.25%2CP',
    'Origin': 'https://investeasy.nipponindiaim.com',
    'Referer': 'https://investeasy.nipponindiaim.com/online/realtime/nav',
    'Sec-Fetch-Dest': 'empty',
    'Sec-Fetch-Mode': 'cors',
    'Sec-Fetch-Site': 'same-origin',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0.0.0 Safari/537.36 Edg/129.0.0.0',
    'X-Requested-With': 'XMLHttpRequest',
    'sec-ch-ua': '"Microsoft Edge";v="129", "Not=A?Brand";v="8", "Chromium";v="129"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"Windows"',
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
