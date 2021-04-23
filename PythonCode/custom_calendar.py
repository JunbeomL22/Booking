# -*- coding: utf-8 -*-
import QuantLib as ql
import utils

# get_holidays returns a list of string, e.g., ['20200415']
kr = utils.get_holidays('KR', days=365*2)
uk = utils.get_holidays('UK', days=365*2)
us = utils.get_holidays('US', days=365*2)
jp = utils.get_holidays('JP', days=365*2)
""" hk = get_holidays('HK', days=365*2) # 현재 hongkong 달력 업데이트 중지 """

#convert them to al.Date
convertor = utils.str2qldate
kr = list(map(convertor, kr))
uk = list(map(convertor, uk))
us = list(map(convertor, us))
jp = list(map(convertor, jp))
""" hk = list(map(convertor, hk)) # 같은 이유로 hk 작업 불 필요 """
# Now, we are ready to go. Define custome_calendar that has the holidays
cKR = ql.SouthKorea()
cUK = ql.UnitedKingdom()
cUS = ql.UnitedStates()
cJP = ql.Japan()
cHK = ql.HongKong()


for d in kr: cKR.addHoliday(d)
for d in uk: cUK.addHoliday(d)
for d in us: cUS.addHoliday(d)
for d in jp: cJP.addHoliday(d)

""" doing nothing for Hong Kong for the time being """

