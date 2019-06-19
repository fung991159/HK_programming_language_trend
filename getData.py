import re, os, sqlite3, time

import requests, bs4, openpyxl
from datetime import datetime

def get_search_result(plang):
    print(f'fetching {plang} search result')
    jobResultRegex = re.compile(r'.*of (\d+) jobs')
    # targetPage = f'https://hk.jobsdb.com/hk/search-jobs/{plang}/1'
    targetPage = f'https://hk.jobsdb.com/hk/jobs/information-technology/1?Key={plang}'
    res = requests.get(targetPage)
    res.raise_for_status()
    soup = bs4.BeautifulSoup(res.text, 'html.parser')
    jobSearchText = soup.find("span", {"data-automation":"searchResultBar"}).getText()
    mo = jobResultRegex.search(jobSearchText)
    return mo.group(1)


def write_to_excel(plang, jobNum):
    print(f'writing {plang} result to xlsx')
    today = datetime.today().date()
    xlsxPath = r'C:\xampp\htdocs\Python\HK_programming_languages_trend\test2.xlsx'
    wb = openpyxl.load_workbook(xlsxPath)
    ws = wb.active
    lastRow = ws.max_row
    currLangList = []
    #get current language column pos
    for i in range(1, ws.max_column+1):
        currLangList.append(ws.cell(1,i).value)

    if plang in currLangList: 
        plangCol = currLangList.index(plang) + 1
    else: #if language is not found, append to the end
        ws.cell(1, ws.max_column+1).value = plang
        plangCol = ws.max_column

    #write date if it latest date isn't today
    try:
        if (ws.cell(lastRow, 1).value).date() != datetime.today().date():  #somehow openpyxl max_row take last blank row instead of last row with content
            ws.cell(lastRow+1, 1).value = today
    except AttributeError:
        ws.cell(lastRow+1, 1).value = today
    
    #write job num if there isn't already a numnber
    if ws.cell(ws.max_row, plangCol).value == None:
        ws.cell(ws.max_row, plangCol).value = int(jobNum)

    wb.save(xlsxPath)

def write_to_db(plang, jobNum):
    conn = sqlite3.connect(r"plang_db.db")
    c = conn.cursor()
    today = datetime.now().strftime('%Y%m%d')

    print(f'writing {plang} result to DB')
    with conn:
        c.execute("INSERT INTO p_lang VALUES (:date, :plang, :count)", 
        {'date': today, 'plang':plang, 'count': jobNum})
    conn.close()

if __name__ == "__main__":
    print('execution start')
    print('wait 60s')
    time.sleep(60)
    # format changed to cater URL link
    plangsDict = {
            'Python' : 'Python', 
            'Java' : 'Java', 
            'Javascript' : 'Javascript', 
            'C++' : 'C%2B%2B', 
            'C#' : 'C%23', 
            'Objective-c' : 'objective-c', 
            'PHP' : 'PHP', 
            'Go' : 'Go', 
            'Swift' : 'Swift', 
            'TypeScript' : 'TypeScript', 
    }
    for plang in plangsDict:
        jobNum = get_search_result(plangsDict[plang])
        try:
            # write_to_excel(plang, jobNum)
            write_to_db(plang, jobNum)
            
        except Exception as e:
            print('error occured, writing to log')
            with open('log.txt', 'a') as f:
                f.write(str(e)+'\n')
            break
            
    print('execution end')
   