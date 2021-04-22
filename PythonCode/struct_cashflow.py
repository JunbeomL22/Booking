import xlwings as xw
import QuantLib as ql
import date_utils
import utils
import payoff

def set_struct_cashflow():
    wb = xw.Book.caller()
    sheet = wb.sheets["StructProduct"]
    rg = sheet.range
    base_info = rg("StructBaseInfo").value
    base_dict = utils.conversion(base_info)
    under_info = rg("StructUnderlying").value
    under_dict = utils.conversion(base_info)
    
    rg("G4:U200").value = None
    rg("W4:X200").value = None
    rg("Z4:AA200").value = None
    # cash_flow period
    issue        = base_dict['Issue_Date']
    mat          = base_dict['Maturity']
    coupon_freq  = base_dict['Coupon_Freq']
    
    fixing_days  = int(base_dict['Fixing_Days'])
    payment_days = int(base_dict['Payment_Days'])

    coupon_conv  = base_dict['Coupon_Conv']
    fixing_conv  = base_dict['Fixing_Conv']
    payment_conv = base_dict['Payment_Conv']
    maturity_conv= base_dict['Maturity_Conv']
    coup_rule = base_dict['Coupon_Schedule']
    cal = base_dict['Calendar1']
    ql_rule = date_utils.date_rule_dict[coup_rule]
    ql_cal = date_utils.cal_dict[cal]
    
    perf_type = base_dict['Type']
    coup = base_dict['Coupon']
    floor = base_dict['Floor']
    cap = base_dict['Cap']
    lev = base_dict['Leverage']
    
    cash_flow = payoff.payoff_factory[base_dict['Prod_Type']](coup, cap, floor, lev, perf_type)

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
    #import pdb;pdb.set_trace()

    if base_dict['Is_Callable'] in ['True', 'TRUE', True] and base_dict['Call_Freq'] not in ['Once', 'Twice', 'Thrice']:
        call_freq = base_dict['Call_Freq']
        call_rule = date_utils.date_rule_dict[base_dict['Call_Schedule']]
        #import pdb;pdb.set_trace()
        call_schedule = list(ql.Schedule(ql.Date.from_date(issue), ql.Date.from_date(mat), ql.Period(call_freq),
                                         ql_cal, ql.Unadjusted, ql.Unadjusted, call_rule, False))
        
        call_cashflows = [[ql.Date.to_date(s)] + [1.0] for s in call_schedule[1:]]
        rg("W4:X200").value = call_cashflows
        
    if base_dict['Is_Puttable'] in ['True', 'TRUE', True] and base_dict['Put_Freq'] not in ['Once', 'Twice', 'Thrice']:
        put_freq = base_dict['Put_Freq']
        put_rule = date_utils.date_rule_dict[base_dict['Put_Schedule']]
        put_schedule = list(ql.Schedule(ql.Date.from_date(issue), ql.Date.from_date(mat), ql.Period(put_freq),
                                        ql_cal, ql.Unadjusted, ql.Unadjusted, put_rule, False))[1:]
        
        put_cashflows = [[ql.Date.to_date(s)] + [1.0] for s in put_schedule[1:]]
        rg("Z4:AA200").value = put_cashflows

    

#xw.Book("BookFile.xlsm").set_mock_caller()
#set_struct_cashflow()
                                
if __name__ == "__main__":
    xw.Book("BookFile.xlsm").set_mock_caller()
    main()







