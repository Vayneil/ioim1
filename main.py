import math
import numpy as np
from openpyxl import load_workbook


def sigma_p(a, e, T, e_dot):
    R = 8.314
    W = math.exp(a[6] * e)
    T1 = math.exp(a[3] / (R * (T + 273)))
    T2 = math.exp(a[5] / (R * (T + 273)))
    e1 = e ** a[1]
    edot1 = e_dot ** a[2]
    result = 1
    result = result * W * a[0] * T1
    result = result + ((1 - W) * a[4] * T2)
    result = result * edot1
    return result


def objective(a, e, T, e_dot, sigma_test):
    sum = 0
    for i in range(len(T)):
        for j in range(len(e[0, :])):
            err_squared = (sigma_test[i, j] - sigma_p(a, e[i, j], T[i], e_dot[i])) ** 2
            err_relative = err_squared / sigma_test[i, j]
            sum = sum + err_relative
    print(sum)
    return sum


def hooke_jeeves(x, s, alpha, epsilon, e, T, e_dot, sigma_test, constraints):
    xb = 0
    while s > epsilon:
        xb = x
        x = trial(xb, s, constraints)
        if objective(x, e, T, e_dot, sigma_test) < objective(xb, e, T, e_dot, sigma_test):
            while objective(x, e, T, e_dot, sigma_test) >= objective(xb, e, T, e_dot, sigma_test):
                xb1 = xb
                xb = x
                x = 2 * xb - xb1
                x = trial(x, s, constraints)
            x = xb
        else:
            s = alpha * s
    return xb


def trial(x, s, constraints):
    for j in range(len(x)):
        trial_x = x
        np.add.at(trial_x, j, s * constraints[1, j])
        if objective(trial_x, e, T, e_dot, sigma_test) < objective(x, e, T, e_dot, sigma_test) and \
                constraints[1, j] > trial_x > constraints[0, j]:
            x = trial_x
        else:
            np.add.at(trial_x, j, -2 * s * constraints[1, j])
            if objective(trial_x, e, T, e_dot, sigma_test) < objective(x, e, T, e_dot, sigma_test) and \
                    constraints[1, j] > trial_x > constraints[0, j]:
                x = trial_x
    return x


# load Excel file and initialize arrays
wb = load_workbook('/home/vay/Desktop/Doswiadczenia.xlsx')
e = []
sigma_test = []
T = []
e_dot = []

# set experimental values
for ws in wb.worksheets:
    e_values = [ws.cell(i, 1).value for i in range(3, 23)]
    sigma_values = [ws.cell(i, 2).value for i in range(3, 23)]
    T.append(ws.cell(2, 4).value)
    e_dot.append(ws.cell(2, 5).value)
    e.append(e_values)
    sigma_test.append(sigma_values)

# convert Python list to NumPy array
sigma_test = np.asarray(sigma_test)
T = np.asarray(T)
e_dot = np.asarray(e_dot)
e = np.asarray(e)
# data = list(zip(e, sigma_test, T, e_dot))
a = np.array([400, 0.5, 0.5, 5000, 0.5, 45000, 0.5])
constraints = np.array([[1, 0, 0, 1, 0, 1, 0], [1000, 1, 1, 10000, 1, 90000, 1]])
# relative step size
s = 0.099
# 0 < alpha < 1
alpha = 0.730
# precision
epsilon = 0.00000001
result = hooke_jeeves(a, s, alpha, epsilon, e, T, e_dot, sigma_test, constraints)
print(result)
