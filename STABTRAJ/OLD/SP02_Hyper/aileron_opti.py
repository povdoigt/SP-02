import xlwings
import numpy as np
from matplotlib import pyplot as plt
from scipy import optimize

b_wb = xlwings.Book('alpha_beta.xlsx')
b_stabilito = b_wb.sheets['Stabilito']

finesse_moy = (b_stabilito.range('G27').value + b_stabilito.range('J27').value) / 2
portance_moy = (b_stabilito.range('G28').value + b_stabilito.range('J28').value) / 2
marge_static_moy = (b_stabilito.range('G29').value + b_stabilito.range('J29').value) / 2
couple_moy = (b_stabilito.range('G30').value + b_stabilito.range('J30').value) / 2

m_bound = [10, 300]
n_bound = [10, 300]
p_bound = [10, 300]
E_bound = [10, 300]


i = 0

def objective_function(x):
    global i
    i += 1
    print(f"Iteration {i}")

    m, n, p, E = x
    b_stabilito.range('C28').value = m
    b_stabilito.range('C29').value = n
    b_stabilito.range('C30').value = p
    b_stabilito.range('C31').value = E
    finesse = b_stabilito.range('H27').value
    portance = b_stabilito.range('H28').value
    marge_static = b_stabilito.range('H29').value
    couple = b_stabilito.range('H30').value
    score = (finesse_moy - finesse) ** 2 + (portance_moy - portance) ** 2 + (marge_static_moy - marge_static) ** 2 + (couple_moy - couple) ** 2
    return score

# res = optimize.shgo(objective_function,
#                     bounds=[m_bound, n_bound, p_bound, E_bound],
#                     n = 10,
#                     options={'disp': True, 'maxfev': 10, 'ftol': 1e-6})

res = optimize.minimize(objective_function,
                        x0=[100, 100, 100, 100],
                        bounds=[m_bound, n_bound, p_bound, E_bound],
                        method='L-BFGS-B',
                        options={'disp': True, 'maxiter': 10, 'ftol': 1e-6})


print(res)