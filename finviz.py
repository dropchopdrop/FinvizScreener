import requests
from bs4 import BeautifulSoup, Comment
import pandas as pd
from tabulate import tabulate
import time
import xlrd
from datetime import date

class GetFinviz():
    def __init__(self,url,today):
        self.fn='ScreenOutStock_'+today+'.xlsx' 
        self.url = url

    def finvizscreener(self):
        df = []
        RemainingRecord = 1
        PageNum = 1

        while RemainingRecord > 0:
            NextPage = self.url + '&r=' + str(PageNum)
            page = requests.get(NextPage, headers = {"User-Agent": "Mozilla/5.0"})
            soup = BeautifulSoup(page.content, 'lxml')
            content = soup.find(id='screener-content')
            if RemainingRecord == 1 and PageNum == 1:
            # table #2 for total record, count page number
                TotalRecord = content.find_all('table')[2] 
                TotalNum = TotalRecord.find_all('td', class_='count-text')[0]
                RemainingRecord = int(TotalNum.text.split()[1])

            # table #3 for potential stock
            table = content.find_all('table')[3] 
            currentpage = pd.read_html(str(table), header=0, index_col=0)
            df.append(currentpage[0])
            RemainingRecord -= 20
            PageNum += 20
            # avoid too frequent request
            time.sleep(10)

        df = pd.concat(df)
        #print( tabulate(df, headers='keys', tablefmt='psql') )
        return df

    def savetoexcel(self,data,sheetname):
        #Save if it contains equities
        if len(data) != 0:
            writer=pd.ExcelWriter(self.fn)
            data.to_excel(writer,sheetname,index=False)
            writer.save()

if __name__ == "__main__":
    url = 'https://finviz.com/screener.ashx?v=111&f=fa_salesqoq_o5,sh_curvol_o50,sh_instown_o10&ft=4'
    today = date.today().strftime("%Y%m%d")
    sample = GetFinviz(url,today)
    Data = sample.finvizscreener()
    if len(Data) != 0:
        sample.savetoexcel(Data,'Sheet1')
    else:
        print('No data return!')
