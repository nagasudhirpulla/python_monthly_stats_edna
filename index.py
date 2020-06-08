# -*- coding: utf-8 -*-
"""
Created on Mon Jun  8 18:10:22 2020

@author: Nagasudhir Pulla

Python script for fetching the monthly stats of lines
"""
import pandas as pd
import calendar
import datetime as dt
from scada_fetcher import fetchPntHistData
import numpy as np

# get 400, 765 kV lines
df400 = pd.read_excel('lines.xlsx', header=None, sheet_name='400')
df765 = pd.read_excel('lines.xlsx', header=None, sheet_name='765')

startDate = dt.datetime(2019,4,1)
endDate = dt.datetime(2020,3,31)
currDate = startDate
# combine to get a single DataFrame
df1 = df765.append(df400, ignore_index=True)
# provision for slicing in case of testing
df = df1.iloc[:, :]

datesIndex = []

maxDataArray = []
minDataArray = []
avgDataArray = []

while currDate <= endDate:   
    datesIndex.append(currDate)
    nextMonthDate = calendar.nextmonth(currDate.year, currDate.month)
    nextMonthDate = dt.datetime(nextMonthDate[0], nextMonthDate[1],1)
    fetchPer = (nextMonthDate-currDate)
    maxDataObj = {}
    minDataObj = {}
    avgDataObj = {}
    for rowIter in range(df.shape[0]):
        pnt = df.iloc[rowIter, 0]
        print('pntNum = {0}, now = {1}, fetchMonth = {2}, pnt = {3}'.format(rowIter, dt.datetime.now(), currDate, pnt))
        # initialize values
        maxDataObj[pnt] = np.nan
        minDataObj[pnt] = np.nan
        avgDataObj[pnt] = np.nan
        
        # fetch max data
        fetchedData = fetchPntHistData(pnt, currDate, nextMonthDate-dt.timedelta(seconds=1), 'max', int(fetchPer.total_seconds()))
        if len(fetchedData) > 0:
            maxDataObj[pnt] = fetchedData[0]['dval']
            
        # fetch min data
        fetchedData = fetchPntHistData(pnt, currDate, nextMonthDate-dt.timedelta(seconds=1), 'min', int(fetchPer.total_seconds()))
        if len(fetchedData) > 0:
            minDataObj[pnt] = fetchedData[0]['dval']
        
        # fetch avg data
        fetchedData = fetchPntHistData(pnt, currDate, nextMonthDate-dt.timedelta(seconds=1), 'average', int(fetchPer.total_seconds()))
        if len(fetchedData) > 0:
            avgDataObj[pnt] = fetchedData[0]['dval']
    maxDataArray.append(maxDataObj)
    minDataArray.append(minDataObj)
    avgDataArray.append(avgDataObj)
    currDate = nextMonthDate

maxDf = pd.DataFrame(maxDataArray, index = datesIndex)
minDf = pd.DataFrame(minDataArray, index = datesIndex)
avgDf = pd.DataFrame(avgDataArray, index = datesIndex)

nowTimeStr = int(dt.datetime.now().timestamp()*1e6)
dumpFilename = 'output/monthly_lines_stats_{}.xlsx'.format(nowTimeStr)

with pd.ExcelWriter(dumpFilename) as writer:
    maxDf.to_excel(writer, sheet_name='max')
    minDf.to_excel(writer, sheet_name='min')
    avgDf.to_excel(writer, sheet_name='avg')
