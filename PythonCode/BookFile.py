import xlwings as xw
import QuantLib as ql
import date_utils

def main():
    wb = xw.Book.caller()
    sheet = wb.sheets["RateProduct"]
    rg = sheet.range
    rg("G4:U200").value = None
    rg("W4:X200").value = None
    rg("Z4:AA200").value = None
    
    # cash_flow period
    issue        = rg("C11").value
    mat          = rg("C12").value
    
    coupon_freq  = rg("C20").value
    
    fixing_days  = int(rg("C28").value)
    payment_days = int(rg("C29").value)

    coupon_conv  = rg("C30").value
    fixing_conv  = rg("C31").value
    payment_conv = rg("C32").value
    maturity_conv= rg("C33").value
    
    rule = rg("C27").value
    cal = rg("C24").value
    
    ql_rule = date_utils.date_rule_dict[rule]
    ql_cal = date_utils.cal_dict[cal]
    #
    perf_type = rg("C34").value
    coup = rg("C35").value
    floor = rg("C36").value
    cap = rg("C37").value
    lev = rg("C38").value
    
    cash_flow = [coup, cap, floor, perf_type, lev, perf_type, -lev]

    schedule = date_utils.coupon_generation(ql.Date.from_date(issue),
                                            ql.Date.from_date(mat),
                                            ql.Period(coupon_freq),
                                            ql_rule,
                                            fixing_days, payment_days, 
                                            ql_cal,
                                            date_utils.conv_dict[coupon_conv],
                                            date_utils.conv_dict[fixing_conv],
                                            date_utils.conv_dict[payment_conv],
                                            date_utils.conv_dict[maturity_conv])
    cash_flows = [s + cash_flow for s in schedule]
    rg("G4:U200").value = cash_flows
    
    # Call Period
    is_call = rg("C18").value
    is_put  = rg("C19").value
    call_freq=rg("C21").value
    put_freq =rg("C22").value
    call_schedule = ql.Schedule(effect
    

xw.Book("BookFile.xlsm").set_mock_caller()
main()
if __name__ == "__main__":
    xw.Book("BookFile.xlsm").set_mock_caller()
    main()







