import QuantLib as ql


effectiveDate = ql.Date(15,6,2020)
terminationDate = ql.Date(15,6,2022)
frequency = ql.Period('6M')
calendar = ql.SouthKorea()
convention = ql.ModifiedFollowing
terminationDateConvention = ql.ModifiedFollowing
rule = ql.DateGeneration.Forward
endOfMonth = False
schedule = ql.Schedule(effectiveDate, terminationDate, frequency, calendar, convention, terminationDateConvention, rule, endOfMonth)


def coupon_generation(effect, term, freq, rule, fixing_days, payment_days,
                      cal = ql.SouthKorea(), 
                      conv =ql.ModifiedFollowing,
                      term_conv=ql.ModifiedFollowing,
                      fixing_conv=ql.ModifiedFollowing,
                      payment_conv=ql.ModifiedFollowing,
                      eom = False):
    schedule = ql.Schedule(effect, term, freq, cal, conv, term_conv, rule, eom)
    
    dates = list(schedule)
    ret = []
    coupon_date = ()
    coupon_length = len(dates)-1
    
    for i in range(coupon_length):
        fixing_date  = cal.advance(schedule[i], ql.Period(-fixing_days, ql.Days), fixing_conv) 
        payment_date = cal.advance(schedule[i+1], ql.Period(payment_days, ql.Days), payment_conv) 
        coupon_date = list(map(ql.Date.to_date, (fixing_date, schedule[i], schedule[i+1], payment_date)))
        ret = ret + [coupon_date]

    return ret
                                  
date_rule_dict ={"Backward": ql.DateGeneration.Backward,
                 "Forward":  ql.DateGeneration.Forward}

cal_dict = {"SouthKorea": ql.SouthKorea(), "Null": ql.NullCalendar()}

conv_dict = {"Following": ql.Following,
             "ModifiedFllowing": ql.ModifiedFollowing,
             "Preceding": ql.Preceding}











