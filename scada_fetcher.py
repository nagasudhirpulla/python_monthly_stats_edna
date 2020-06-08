# -*- coding: utf-8 -*-
"""
Created on Mon Jun  8 18:10:22 2020

@author: Nagasudhir Pulla
"""
import requests
import json
import datetime as dt

def makeTwoDigits(num):
    if(num<10):
        return "0"+str(num)
    return num
    
def fetchPntHistData(pnt, startTime, endTime, fetchStrategy='snap', secs=300):
    startTimeStr = startTime.strftime('%d/%m/%Y/%H:%M:%S')
    endTimeStr = endTime.strftime('%d/%m/%Y/%H:%M:%S')
    # print(req_date_str)
    params = dict(
            pnt=pnt,
            strtime =  startTimeStr,
            endtime =  endTimeStr,
            secs = secs,
            type=fetchStrategy
            )
    # http://wmrm0mc1:62448/api/values/history?pnt=WRLDCMP.SCADA1.A0001347&strtime=12/12/2019/00:00:00&endtime=13/12/2019/00:00:00&secs=900&type=average
    r = requests.get(url = "http://wmrm0mc1:62448/api/values/history", params = params)
    data = json.loads(r.text)
    return data