import numpy as np
import xlwings as xw
import matplotlib.pyplot as plt

@xw.func
def test_func(x, y):
    return 2*(x+y)

@xw.func
@xw.arg('x', np.array, ndim=2)
@xw.arg('y', np.array, ndim=2)
def matrix_mult(x, y):
    return x @ y

@xw.func
def myplot(n, caller):
    import pdb;pdb.set_trace()
    caller.sheet.range("A1").value = 1.0


if __name__ == '__main__':
    xw.serve()
