import pandas as pd


def data_munging(category,data_num):
    '''
    每個產業的資料清理
    參數：產業代號(category)--數字
         處理資料類(data_num)--數字
            3 -- 資產負債表
            4 -- 綜合損益表
            5 -- 現金流量表
            6 -- 每月盈餘
            7 -- 每月股價
    '''
    cate = category[i]
    schedule = list(map(int,list(index.iloc[:,i].dropna())))
    index = []
    for i,company in enumerate(schedule):
        data = pd.read_excel('Desktop/ntumath/'+cate+'/'+str(company)+'/'+str(company)+'(102.1~105.4)EPM.xlsx')
        index.append(data.index)
        name = set(index)#把它拆開


    for item in name:
        for subitem in index:
            if item in subitem:
                count(item) += 1;
            else:
                pass

    for subset in set(index):
        if count(subset)/len(index) < 0.7
            del(sebset)
        else:
            pass

    
