from collections import Counter
def check_crawl(category_name,schedule):
    '''
    檢查每個公司資料夾是否只有六個檔案，只回傳True or False
    參數：(產業代號--字串,產業裡的公司--數字)
    '''
    for i,company in enumerate(schedule):
        number_of_files = len(os.listdir('Desktop/ntumath/computation/'+category_name+'/'+str(company)+'/'))
        if number_of_files != 6:
            return False
        else:
            pass
    return True

def data_munging(category_num,data_num):
    '''
    每個產業的資料清理
    參數：產業代號(category)--數字
         處理資料類(data_num)--數字
            3 -- 資產負債表
            4 -- 綜合損益表
            5 -- 現金流量表
    '''
    cate = category[category_num]
    schedule = list(map(int,list(index.iloc[:,category_num].dropna())))
    index_collect = []
    index_set = set()
    if check_crawl(cate,schedule) == False:
        return 'something is wrong'
    number_of_company = len(schedule)
    for i,company in enumerate(schedule):
        combine(cate,company,data_num)
        data = pd.read_csv('Desktop/ntumath/computation/'+cate+'/'+str(company)+'/'+
                            os.listdir('Desktop/ntumath/computation/'+cate+'/'+str(company)+'/')[data_num - 1],index_col=0)
        for row in list(data.index):
            index_collect.append(row)
    counts = Counter(index_collect)
    for index_in_counts in dict(counts):
        if counts[index_in_counts]/number_of_company >= 0.77:
            index_set.add(index_in_counts)
        else:
            pass
    for company in schedule:
        data = pd.read_csv('Desktop/ntumath/computation/'+cate+'/'+str(company)+'/'+
                    os.listdir('Desktop/ntumath/computation/'+cate+'/'+str(company)+'/')[data_num - 1],index_col=0)
        oringin = len(data.index)
        for row_check in list(data.index):
            if row_check not in index_set:
                data.drop(row_check,inplace=True)
            else:
                pass
        after = len(data.index)
        data.to_csv('Desktop/ntumath/computation/'+cate+'/'+str(company)+'/'+
                    os.listdir('Desktop/ntumath/computation/'+cate+'/'+str(company)+'/')[data_num - 1])
        print(str(company)+':')
        print('\t oringin:'+str(oringin))
        print('\t after:'+str(after))
