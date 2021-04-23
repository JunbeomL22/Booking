# -*- coding: utf-8 -*-
import xlwings as xw
import pandas as pd
import utils
import numpy as np
from sqlalchemy import exc
from sqlalchemy.orm import sessionmaker
import database_type as dtp
import time
import custom_calendar
import QuantLib as ql
import pdb
import updater
from win32com.client import pythoncom
@xw.func
def ir_vol_update():
    """
    From BloombergDate.xlsm, retrieve swaption and cap vol.
    Then, it inserts into database FICC_SWAPTION_ATM_VOL and FICC_CAP_ATM_VOL
    """
    wb = xw.Book.caller()
    ws = wb.sheets("IR_VOL")

    # Declare connection object    
    swaption_vol = ws.range("SwaptionVol").value
    cap_vol = ws.range("CapVol").value
    engine = utils.db_engine(database = 'otcora', schema = 'OTCUSER', password = 'otcuser')
    Session = sessionmaker(bind=engine)
    session = Session()
    # The code below inputs swaption vol data
    updater.updater(data = swaption_vol, table_name = 'ficc_swaption_atm' ,
                    head_nullable_data=4, date_index = 0, factor = 0.0001,
                    data_name = 'swaption vol/premium data',
                    engine = engine, session = session)
    # The code below inputs cap vol data
    updater.updater(data = cap_vol, table_name = 'ficc_cap_atm',
                    head_nullable_data=3, date_index = 0, factor = 0.0001,
                    data_name = 'cap vol/premium data',
                    engine = engine, session = session)
    session.close()
    engine.dispose() 
    
    #utils.Mbox("", "swaption & cap vol done", 0)

def ir_vol_history_update(ws=None):
    """
    ir_vol_his_update()
    From BloombergDate.xlsm, retrieve swaption and cap vol data
    from start_dt to end_dt.
    """
    wb = xw.Book.caller()
    ws = wb.sheets("IR_VOL")
    #
    period_range = "ir_vol_period"
    date_range = "ir_vol_date"
    # To pause for collecting data from bloomberg
    time_to_sleep = 2 # seconds
    # get range of date
    st_dt, end_dt = ws.range( period_range ).value
    date = utils.str2qldate(st_dt)
    end_dt = utils.str2qldate(end_dt)
    day1 = ql.Period(1, ql.Days)
    # change ql.Date to str
    str_converter = utils.qldate2str
    # calendar to deal with holidays
    cKR = custom_calendar.cKR
    # DB connection engine
    engine = utils.db_engine(database = 'otcora', schema = 'OTCUSER', password = 'otcuser')
    Session = sessionmaker(bind=engine)
    session = Session()
    while date <= end_dt:
        if cKR.isBusinessDay(date):
            ws.range( date_range ).value = str_converter(date)
            
            checking = utils.check_bloomberg_error
            checker = 0
            swaption_vol = ws.range("SwaptionVol").value
            cap_vol = ws.range("CapVol").value
            while checking(swaption_vol, time_to_sleep):
                checker = checker + 1
                if checker > 3:
                    raise utils.JbException("hello")
            """
                
                pythoncom.PumpWaitingMessages()
                time.sleep(0.01)
                
            """
            # 
            updater.updater(data = swaption_vol, table_name = 'ficc_swaption_atm_vol' ,
                            head_nullable_data=4, date_index = 0, factor = 0.0001,
                            data_name = 'swaption vol data',
                            engine = engine, session = session)
            # 
            updater.updater(data = cap_vol, table_name = 'ficc_cap_atm_vol',
                            head_nullable_data=3, date_index = 0, factor = 0.0001,
                            data_name = 'cap vol data',
                            engine = engine, session = session)
        date = date + day1

    #ms = wb.app.api()
    # from xlwings.constants import Calculation
    # xw.constants.CalculationState
    #
    #

    
    session.close()
    engine.dispose()
    
    utils.Mbox("", "swaption & cap vol done", 0)
    
#xw.Book('../BloombergData.xlsm').set_mock_caller()
#ir_vol_history_update()

