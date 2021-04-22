def powerspread_cashflow(day_count, coup, cap, floor, lev, perf_type='Ave', length=-1):
    return [day_count, coup, cap, floor, perf_type, lev, -lev]

def els_cashflow(day_count, coup, cap, floor, lev, perf_type='Ave', length=-1):
    return [day_count, coup, perf_type]

payoff_factory = {"Power Spread": powerspread_cashflow,
                  "ELS": els_cashflow,
                  "Lizard ELS": els_cashflow}











