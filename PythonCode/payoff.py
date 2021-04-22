def powerspread_cashflow(coup, cap, floor, lev, perf_type='Ave', length=-1):
    return [coup, cap, floor, perf_type, lev, perf_type, -lev]

payoff_factory = {"Power Spread": powerspread_cashflow}










