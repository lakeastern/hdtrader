# -*- coding: utf-8 -*-
import pandas as pd
import numpy as np
import requests
import bs4
import urllib.parse
import matplotlib.pyplot as plt
import numpy

# 전 종목 코드 구해 오기
def download_stock_codes(market=None, delisted=False):
    
    MARKET_CODE_DICT = {
    'kospi': 'stockMkt',
    'kosdaq': 'kosdaqMkt',
    'konex': 'konexMkt'
    }

    DOWNLOAD_URL = 'kind.krx.co.kr/corpgeneral/corpList.do'
    
    params = {'method': 'download'}

    if market.lower() in MARKET_CODE_DICT:
        params['marketType'] = MARKET_CODE_DICT[market]

    if not delisted:
        params['searchType'] = 13

    params_string = urllib.parse.urlencode(params)
    request_url = urllib.parse.urlunsplit(['http', DOWNLOAD_URL, '', params_string, ''])

    df = pd.read_html(request_url, header=0)[0]
    df.종목코드 = df.종목코드.map('{:06d}'.format)

    return df

# 시가총액 데이터를 가져와 데이터프레임으로 만드는 함수
def make_cap_dataframe(firm_code):
    cap_url = 'https://comp.fnguide.com/SVO2/asp/SVD_Main.asp?pGB=1&cID=&MenuYn=Y&ReportGB=D&NewMenuID=103&stkGb=701&gicode=' + firm_code
    cap_page = requests.get(cap_url)
    cap_tables = pd.read_html(cap_page.text)

    temp_df = cap_tables[0]
    temp_df = temp_df.set_index(temp_df.columns[0])
    temp_df = temp_df[temp_df.columns[:1]]

    if numpy.isnan(float(temp_df[1]['시가총액(보통주,억원)'])):
        temp_df = temp_df.loc[['종가/ 전일대비', '시가총액(상장예정포함,억원)', '발행주식수(보통주/ 우선주)']]
        temp_df.index = ['종가','시가총액','발행주식수']
    else:
        temp_df = temp_df.loc[['종가/ 전일대비', '시가총액(보통주,억원)', '발행주식수(보통주/ 우선주)']]
        temp_df.index = ['종가','시가총액','발행주식수']

    for id in temp_df.index:
        temp_df.loc[id] = temp_df.replace(np.nan, '', regex=True).loc[id].str.split('/')[1][0]
        
    temp_df = temp_df.T
    temp_df.index = [firm_code]
    cap_df = temp_df
    
    return cap_df


def make_fsale_dataframe(firm_code, temp_df):
    fs_url = 'https://comp.fnguide.com/SVO2/asp/SVD_Main.asp?pGB=1&cID=&MenuYn=Y&ReportGB=&NewMenuID=103&stkGb=701&gicode=' + firm_code
    fs_page = requests.get(fs_url)
    fs_tables = pd.read_html(fs_page.text)
    
    fsale_df = fs_tables[12]
    fsale_df = fsale_df[fsale_df.columns[[0,2,3,4,5]]]
    fsale_df = fsale_df.set_index(fsale_df.columns[0])
    fsale_df.columns = ['']* len(fsale_df.columns)    
    fsale_df.columns = temp_df.columns
    fsale_df = fsale_df.iloc[[0]]
    fsale_df = pd.concat([fsale_df, temp_df.loc[['당기순이익']]])
    return fsale_df  


# [코드 3.15] 재무제표 데이터를 가져와 데이터프레임으로 만드는 함수 (CH3. 데이터 수집하기.ipynb)

def make_fs_dataframe(firm_code):
    
    try:

        fs_url = 'https://comp.fnguide.com/SVO2/asp/SVD_Finance.asp?pGB=1&cID=&MenuYn=Y&ReportGB=&NewMenuID=103&stkGb=701&gicode=' + firm_code
        fs_page = requests.get(fs_url)
        fs_tables = pd.read_html(fs_page.text)

        temp_df = fs_tables[1]
        temp_df = temp_df.set_index(temp_df.columns[0])
        temp_df = temp_df[temp_df.columns[:4]]
        if temp_df.index[0] == '매출액':
            temp_df = temp_df.loc[['매출액', '당기순이익']]
        else:
            temp_df = make_fsale_dataframe(firm_code, temp_df)

        temp_df2 = fs_tables[3]
        temp_df2 = temp_df2.set_index(temp_df2.columns[0])
        temp_df2 = temp_df2.loc[['자본']]

        temp_df3 = fs_tables[5]
        temp_df3 = temp_df3.set_index(temp_df3.columns[0])
        temp_df3 = temp_df3.loc[['영업활동으로인한현금흐름']]


        fs_df = pd.concat([temp_df, temp_df2, temp_df3])
        fs_df.index = ['매출액','순이익','자본','영업현금흐름']
        
        fs_df = pd.DataFrame({firm_code:[fs_df.columns[3],
                                 fs_df.loc['매출액'].sum(),
                                 fs_df.loc['순이익'].sum(),
                                 fs_df.loc['자본'].dropna()[-1],
                                 fs_df.loc['영업현금흐름'].sum()]},
            index=['날짜','매출액','순이익','자본','영업현금흐름'])

        
        fs_df = fs_df.T
    except:
        pass
    
    return fs_df


def make_fhd_dataframe(firm_code):
    
    try:
        fhd_url = 'https://comp.fnguide.com/SVO2/asp/SVD_Main.asp?pGB=1&cID=&MenuYn=Y&ReportGB=D&NewMenuID=103&stkGb=701&gicode=' + firm_code
        fhd_page = requests.get(fhd_url)
        fhd_tables = pd.read_html(fhd_page.text)
        temp_df = fhd_tables[11]
        temp_df = temp_df.set_index(temp_df.columns[0])
        fhd_df = pd.DataFrame({firm_code:[temp_df.loc['BPS(원)'][4],
                                          temp_df.loc['ROE'][5:].mean(),
                                          temp_df.loc['부채비율'][5:].mean(),
                                          temp_df.loc['배당수익률'][2:5].mean()]},
                              index=['BPS','ROE3년','부채비율3년','배당3년'])
        fhd_df = fhd_df.T
        

    except:
        fhd_df = pd.DataFrame({firm_code:[np.nan,
                                          np.nan,
                                          np.nan,
                                          np.nan]},
                              index=['BPS','ROE3년','부채비율3년','배당3년'])
        fhd_df = fhd_df.T
    
    
    return fhd_df
    



def get_value_rank(sorted_cap_value, value_type):
    sorted_cap_value[value_type] = pd.to_numeric(sorted_cap_value[value_type])
    value_sorted = sorted_cap_value.sort_values(by = value_type, ascending=False)
    value_sorted[value_type + '순위'] = value_sorted[value_type].rank(ascending=False)    
    return value_sorted


def get_small_cap(value_df, ratio=[0.8, 1.0]):   
    temp_df = value_df.set_index('종목코드')
    temp_df['시가총액'] = pd.to_numeric(temp_df['시가총액'])
    temp_df = temp_df.sort_values(by='시가총액', ascending=False)
    temp_df['시가총액비율'] = temp_df['시가총액'].rank(ascending=False) / temp_df.shape[0]
    sorted_cap_value = temp_df[(temp_df['시가총액비율'] > min(ratio)) & (temp_df['시가총액비율'] <= max(ratio))]
    return sorted_cap_value  


def value_combo(sorted_cap_value, value_list, num=50):
    for i, value in enumerate(value_list):
        if i == 0:            
            value_combo = get_value_rank(sorted_cap_value, value)
            value_combo['combo_rank'] = value_combo[value + '순위']     
            value_combo = value_combo[value_combo[value] > 0]
        else:            
            value_combo = get_value_rank(value_combo, value)
            value_combo['combo_rank'] = value_combo['combo_rank'] + value_combo[value + '순위']
            value_combo = value_combo[value_combo[value] > 0]
            
    value_combo = value_combo[value_combo['1/PBR'] <= 5]
    value_combo = value_combo.sort_values(by='combo_rank')
    value_combo['combo_rank'] = value_combo['combo_rank'].rank()
    value_combo = value_combo[['기업명','종가','시가총액','combo_rank','날짜'] + ['{0}순위'.format(i) for i in value_list] + value_list]
    value_combo = value_combo[:num]
    return value_combo