#거래량 급증 종목 검색
# 60일 평균 거래량 대비 10배이상 거래량 급증 종목 출력
# 대신증권 Cybos plus API 이용

import win32com.client

def CheckVolumn(instStockChart,code):

    # set input value
    instStockChart.SetInputValue(0,code)
    instStockChart.SetInputValue(1,ord('2'))
    instStockChart.SetInputValue(4,60)
    instStockChart.SetInputValue(5,8)
    instStockChart.SetInputValue(6,ord('D'))
    instStockChart.SetInputValue(9,ord('1'))

    # Block Request
    instStockChart.BlockRequest()

    # Get Data
    volumes =[]
    numData = instStockChart.GetHeaderValue(3)
    for i in range(numData):
        volume = instStockChart.GetDataValue(0,i)
        volumes.append(volume)



    # Calulate average volume
    averageVolume = (sum(volumes)-volumes[0])/(len(volumes)-1)

    if(volumes[0]>averageVolume * 10):
        return 1
    else:
        return 0

if __name__ =='__main__':
    instCpCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr')
    instStockChart = win32com.client.Dispatch('CpSysDib.StockChart')
    codeList = instCpCodeMgr.GetStockListByMarket(1)

    buyList=[]
    print("now it is working")

for code in codeList:
    if CheckVolumn(instStockChart,code) == 1:
        buyList.append(code)
        name = instCpCodeMgr.CodeToName(code)
        print(code,name)

