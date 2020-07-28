import numpy as np
import scipy.optimize as optimize
import matplotlib.pyplot as plt

np.set_printoptions(formatter={"float_kind": lambda x: "%g" % x})

sample_weight = 50  # g
stucco_phases = [2.5 / 100, 85.0 / 100, 2.5 / 100, 10.0 / 100]  # %
dh_sample_purity = 1.0
hh_sample_purity = 0.9
a3_sample_purity = 0.9
hh_min = 75.0
hh_max = 90.0
dh_max = 10.0
a3_max = 5.0
ine_max = 10.0
step = 0.7
rmse_threshold = 1

dh_g_max = None
hh_g_max = 10
a3_g_max = None
in_g_max = None

def best_solution(p):
    """
    this function calculated best solution for weight of each sample (dh, hh, a3, stucco, inert) based of given ideal
    phases.
    :param p: 4x1 numpy array containing weight percents of phases:
                p[0]: weight percent of dh phase
                p[1]: weight percent of hh phase
                p[2]: weight percent of a3 phase
                p[3]: weight percent of inert phase
    :return: w, 5x1 numpy array containing weight of each sample (dh, hh, a3, stucco, inert)
    """

    d, h, a, i = p
    mat_A = np.zeros((4, 5))
    b = np.zeros(4)

    mat_A[0, 0] = d / 100 - dh_sample_purity
    mat_A[0, 1] = d / 100
    mat_A[0, 2] = d / 100
    mat_A[0, 3] = d / 100 - stucco_phases[0]
    mat_A[0, 4] = d / 100

    mat_A[1, 0] = h / 100
    mat_A[1, 1] = h / 100 - hh_sample_purity
    mat_A[1, 2] = h / 100
    mat_A[1, 3] = h / 100 - stucco_phases[1]
    mat_A[1, 4] = h / 100

    mat_A[2, 0] = a / 100
    mat_A[2, 1] = a / 100
    mat_A[2, 2] = a / 100 - a3_sample_purity
    mat_A[2, 3] = a / 100 - stucco_phases[2]
    mat_A[2, 4] = a / 100

    mat_A[3, 0] = i / 100 - 1 + dh_sample_purity
    mat_A[3, 1] = i / 100 - 1 + hh_sample_purity
    mat_A[3, 2] = i / 100 - 1 + a3_sample_purity
    mat_A[3, 3] = i / 100 - stucco_phases[3]
    mat_A[3, 4] = i / 100 - 1

    # p = np.linalg.solve(mat_A, b)

    # lsd = optimize.lsq_linear(mat_A, b, bounds=(lbond, ubond), verbose=0)
    # p = lsd.x

    f = lambda x: np.dot(np.dot(mat_A, x) - b, np.dot(mat_A, x) - b)
    cons = ({'type': 'eq', 'fun': lambda x: x.sum() - sample_weight})
    # bnds = ((0, 10), (0, 50), (0, 10), (0, None), (0, None))
    bnds = ((0, dh_g_max), (0, hh_g_max), (0, a3_g_max), (0, None), (0, in_g_max))
    res = optimize.minimize(f, [0, 0, 0, 0, 0], method='SLSQP', bounds=bnds, constraints=cons,
                            options={'disp': False})
    w = res['x']

    return w


def phase_from_weights(w):
    """
    this function calculated weight percent of each phase (dh, hh, a3, inert) based on weight of each constituent used
    to build the sample.
    :param w: 5x1 numpy array containing weight of each constituent:
                w[0] = weight of dh sample (purity = dh_sample_purity)
                w[1] = weight of hh sample (purity = hh_sample_purity)
                w[2] = weight of a3 sample (purity = a3_sample_purity)
                w[3] = weight of stucco sample (phase contents = stucco_phases)
                w[4] = weight of inert sample
    :return: p, 4x1 numpy array containing phase contents resulted from the combinations of above weights
    """

    p = np.zeros(4)
    w_sum = np.sum(w)
    p[0] = 100 * (w[0] * dh_sample_purity + w[3] * stucco_phases[0]) / w_sum
    p[1] = 100 * (w[1] * hh_sample_purity + w[3] * stucco_phases[1]) / w_sum
    p[2] = 100 * (w[2] * a3_sample_purity + w[3] * stucco_phases[2]) / w_sum
    p[3] = 100 - p[0] - p[1] - p[2]

    return p


# lbond = np.array([0, 0, 0, 0.7 * sample_weight, 0])  # lower limit of dhs, hhs, a3s, stucco, ines
# ubond = np.array([20, 20, 20, sample_weight, 20])  # lower limit of dhs, hhs, a3s, stucco, ines
ideal_sample_num = 0
actual_sample_num = 0
ideal_phases = []
actual_phases = []
weights = []
rmse = []

for hi in range(int(hh_min / step), int(hh_max / step) + 1):
    for di in range(0, int(dh_max / step) + 1):
        for ai in range(0, int(a3_max / step) + 1):
            for ii in range(0, int(ine_max / step) + 1):
                # ideal phase contents
                h = hi * step
                d = di * step
                a = ai * step
                i = ii * step
                if int(h + d + a + i) == 100:
                    # finding the best solution for weight of each sample that give the ideal phase contents
                    ideal_sample_num += 1
                    print('ideal sample #: ', ideal_sample_num)
                    p_ideal = [d, h, a, i]

                    w = best_solution(p_ideal)

                    p = phase_from_weights(w)

                    temp_rmse = np.sqrt(np.mean(np.square(np.subtract(p_ideal, p))))

                    if temp_rmse <= rmse_threshold:
                        actual_sample_num += 1
                        # print('Ideal phases: ', p_ideal)
                        # print('Actual phases: ', np.round(p, 2))
                        # print('sample weights: ', np.round(w, 2))
                        ideal_phases.append(p_ideal)
                        actual_phases.append(p)
                        weights.append(w)
                        rmse.append(temp_rmse)

print('actual sample #: ', actual_sample_num)

# ideal_phases = np.array(ideal_phases)
# actual_phases = np.array(actual_phases)
#
# rmse = np.sqrt(np.mean(np.square(np.subtract(ideal_phases, actual_phases)), axis=1))
x = range(len(ideal_phases))
pa = []
xlabels = ['DH (%)', 'HH (%)', 'AIII (%)', 'IN (%)']
ylabels = ['RMSE (%)']
plt.figure()
for i in range(4):
    pa_i = [e[i] for e in ideal_phases]
    pa_a = [e[i] for e in actual_phases]
    plt.subplot(2, 2, i + 1)
    plt.scatter(pa_i, rmse, color='blue')
    plt.scatter(pa_a, rmse, color='red')
    plt.xlabel(xlabels[i], fontsize=18)
    plt.ylabel('RMSE', fontsize=18)

plt.show()

costs = np.array([0.5360, 0.0599, 0.0599, 0, 0])
total_weights = np.sum(weights, axis=0)
print(total_weights, np.dot(costs, total_weights.T))
