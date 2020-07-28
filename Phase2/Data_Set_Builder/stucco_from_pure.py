import numpy as np
import scipy.optimize as optimize

lbond = np.zeros(4)
ubond = lbond + 1
pure_samples = [[100.0, 0.0, 0.0, 0.0], [1.015, 97.132, 1.853, 0.0], [1.015, 0.0, 98.985, 0.0], [0.0, 0.0, 0.0, 100.0]]
HH_AIII = [[100.0, 0.0, 0.0, 0.0], [1.98365, 88.26742, 0.13744, 9.61149], [0.5056, 7.71697, 68.77165, 23.00579]]
HHP_AIII = [[100.0, 0.0, 0.0, 0.0], [1.015, 97.132, 1.853, 0.0], [0.5056, 7.71697, 68.77165, 23.00579]]
HH_AIIIP = [[100.0, 0.0, 0.0, 0.0], [1.98365, 88.26742, 0.13744, 9.61149], [1.015, 0.0, 98.985, 0.0]]

stucco = [5, 75, 10, 10]

def exact_sol(samples, stucco):
    A = np.zeros((4, 4))
    b = np.zeros(4)

    for i in range(4):
        b[i] = stucco[i]/100
        for j in range(4):
            A[i][j] = samples[j][i]/100

    w = list(np.linalg.solve(A, b))
    # f = lambda x: np.dot(np.dot(A, x) - b, np.dot(A, x) - b)
    # cons = ({'type': 'eq', 'fun': lambda x: x.sum() - 1})
    # bnds = ((0, 1), (0, 1), (0, 1), (0, 1))
    # res = optimize.minimize(f, [0, 0, 0, 0], method='SLSQP', bounds=bnds, constraints=cons,
    #                         options={'disp': False})
    # w = res['x']

    return w

def const_sol(samples, stucco):
    A = np.zeros((4, 3))
    b = np.zeros(4)

    for i in range(4):
        b[i] = stucco[i]/100
        for j in range(3):
            A[i][j] = samples[j][i]/100

    f = lambda x: np.dot(np.dot(A, x) - b, np.dot(A, x) - b)
    cons = ({'type': 'eq', 'fun': lambda x: x.sum() - 1})
    bnds = ((0, 1), (0, 1), (0, 1))
    res = optimize.minimize(f, [0, 0, 0], method='SLSQP', bounds=bnds, constraints=cons,
                            options={'disp': False})
    w = res['x']

    return w



comp = const_sol(HH_AIIIP, stucco)
comp = np.around(comp, decimals=4)
s1 = np.zeros(4)
for i in range(3):
    s1 = s1 + np.dot(HH_AIIIP[i], comp[i])

print('---------Constraint Solution----------')
print('weights of each sample in stucco')
comp = np.multiply(comp, 100.0)
print(comp)

print('target stucco phase content:')
print(stucco)

print('calculated stucco phase content:')
print(np.around(s1, decimals=2))
######################################################
comp = exact_sol(pure_samples, stucco)
comp = np.around(comp, decimals=4)
s1 = np.zeros(4)
for i in range(4):
    s1 = s1 + np.dot(pure_samples[i], comp[i])

print('---------Exact Solution----------')
print('weights of each sample in stucco')
comp = np.multiply(comp, 100.0)
print(comp)

print('target stucco phase content:')
print(stucco)

print('calculated stucco phase content:')
print(np.around(s1, decimals=2))




