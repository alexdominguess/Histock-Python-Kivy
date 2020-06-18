import requests
from bs4 import BeautifulSoup as soup 
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import os, sys
import time

def open_chrome(stock, frequency):
    stock = stock.upper()
    if frequency == 'Mensal':
        url = 'https://br.financas.yahoo.com/quote/'+stock+'.SA/history?period1=0&period2=9999999999&interval=1mo&filter=history&frequency=1mo'
    elif frequency == 'Semanal':
        url = 'https://br.financas.yahoo.com/quote/'+stock+'.SA/history?period1=0&period2=9999999999&interval=1wk&filter=history&frequency=1wk'
    else:   
        url = 'https://br.financas.yahoo.com/quote/'+stock+'.SA/history?period1=0&period2=9999999999&interval=1d&filter=history&frequency=1d'

    scriptpath = os.path.dirname(sys.argv[0])
    driver = webdriver.Chrome(scriptpath + '\\chromedriver.exe')
    #driver.set_window_position(-10000,0)
    driver.get(url)
    return driver


def read_yahoo_page(driver):
    html = driver.find_element_by_tag_name('html')
    connected = html.text[:12]
    if connected == 'Sem Internet':
        return "Sem Internet"
    else:
        return html

#Loading data...
def get_stocks(html, driver):
    while True:
        html = driver.find_element_by_tag_name('html')
        html.send_keys(Keys.END)
        driver.execute_script("return document.documentElement.outerHTML")
        html = driver.page_source
        html = soup(html,"html.parser")
        spans = html.findAll("span")
        spans.reverse()
        #check if there is more info to load
        action = ""
        for span in spans:
            get_text = span.text
            texts = get_text.split()
            if 'Carregando mais dados...' in span.text:###################################
                action = "carregar"
                break
        if action == "carregar":
            pass
        else:
            break

    #After load all page, check is stock exist
    html = driver.page_source
    driver.quit()
    html = soup(html,"html.parser")
    spans = html.findAll("span")
    #check if stock exist
    for span in spans:
        get_text = span.text
        texts = get_text.split()
        for text in texts:
            if text == 'Nenhum':
                return 'Nenhum'

    #if it exist add line by line in a list of dictionaries    
    tables = html.findAll("table")
    table = tables[0]
    rows = table.findAll("tr")
    history_list = []
    info_type = ['Data', 'Abertura', 'Maxima', 'Minima', 'Fechamento', 'Fechamento Ajustado', 'Volume']
    rows_qty = len(rows)
    for row in range (rows_qty):
        columns = rows[row]
        columns = columns.findAll("td")
        values_dict  = {}
        for column in range(len(columns)):
            value = columns[column]
            value = value.text
            values_dict[info_type[column]] = value
        
        if len(columns) > 0:
            history_list.append(values_dict)
    return history_list



def save_results(history_data, stock):
    scriptpath = os.path.dirname(sys.argv[0])
    history_cotation_df = pd.DataFrame(history_data)
    history_cotation_df.to_excel(scriptpath + '\\results.xls', sheet_name=stock, index=False)


