#업종별 평균 PER 보다 작은 종목 검색 및
# 종목 코드, 종목명 리스트 csv 파일 작성

import win32com.client

instCpCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr')
instMarketEye = win32com.client.Dispatch('CpSysDib.MarketEye')
issue_code = input("업종코드를 입력하세요 :")
tarketCodeList = instCpCodeMgr.GetGroupCodeList(issue_code)

# Get PER
instMarketEye.SetInputValue(0,(0,67))
instMarketEye.SetInputValue(1,tarketCodeList)

# Block Request
instMarketEye.BlockRequest()


#
# Get HeaderValue
numStock = instMarketEye.GetHeaderValue(2)

# Get data
sumPer = 0
code_list =[]
for i in range(numStock):
    sumPer += instMarketEye.GetDataValue(1,i)
    code_list.append(instMarketEye.GetDataValue(0,i))

# calculate average PER
avgPer = sumPer/numStock
print("Average PER:",avgPer)

# open csv file
f = open("low_per.csv",'w')
for i in range(numStock):
    if instMarketEye.GetDataValue(1,i) < avgPer and instMarketEye.GetDataValue(1,i) !=0: # compare with avg PER and exclude PER=0
        print(code_list[i],instCpCodeMgr.CodeToName(code_list[i]),instMarketEye.GetDataValue(1,i))
        f.write("%s, %s ,%s\n" % (code_list[i],instCpCodeMgr.CodeToName(code_list[i]),instMarketEye.GetDataValue(1,i)))
f.close()





