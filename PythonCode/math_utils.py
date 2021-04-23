from scipy.interpolate import interp1d

class FlatExtrapolateLinear:
    def __init__(self, x, y):
        self.lin = interp1d(x, y, bounds_error = False, fill_value = y[-1])
        self.leftx = x[0]
        self.lefty = y[0]

    def __call__(self, u):
        if u <= self.leftx:
            return self.lefty
        else:
            return self.lin(u)
    
