import xlrd
import xlwt
import numpy as np
from scipy.optimize import lsq_linear
from pathlib import Path


w0 = 136.14

DW = {
    "dhwl": 2 * 18.01528 / (w0 + 2 * 18.01528),
    "hhwl": 0.5 * 18.01528 / (w0 + 0.5 * 18.01528),
    "hhwg": 1.5 * 18.01528 / (w0 + 0.5 * 18.01528),
    "a3wg1": 0.5 * 18.01528 / w0,
    "a3wg2": 2 * 18.01528 / w0
}


def isdigit(str0):
    str1 = str0.replace('.', '', 1)
    str2 = str1.replace('-', '', 1)
    str3 = str2.replace('+', '', 1)
    return str3.isdigit()


def isListEmpty(inList):
    if isinstance(inList, list): # Is a list
        return all(map(isListEmpty, inList))
    return False # Not a list


def gpa_linear(wl_id, dwl, dwa, dwh):
    """
    Function to calculate components of system of linear equations giving stucco phase content from two weight gain
    and one weight loss measurements.
    :param wl_id: The id of the weight loss measurement. 1: ORG sample, 2: HUM sample, 3: HYD sample
    :param dwl: weight loss wl_id sample
    :param dwa: weight gain of HUM sample
    :param dwh: weight gain of HYD sample
    :return:
    A: 4x4 matrix containing coefficients of system of linear equations giving X=[DH, HH, AIII, FM, IN]
    b: array of length 4 giving the constants in AX - b = 0 equation.
    """
    A = np.zeros((4, 4))
    b = np.zeros(4)
    if dwa >= 0:
        A[0, 0] = 0
        A[0, 1] = 0
        A[0, 2] = DW["a3wg1"]
        A[0, 3] = 0
        A[1, 0] = 0
        A[1, 1] = DW["hhwg"]
        A[1, 2] = DW["a3wg2"]
        A[1, 3] = 0
        A[3, 0] = 1
        A[3, 1] = 1
        A[3, 2] = 1
        A[3, 3] = 1
        if wl_id == 0:
            A[2, 0] = DW["dhwl"]
            A[2, 1] = DW["hhwl"]
            A[2, 2] = 0
            A[2, 3] = 0
            b[2] = dwl
        if wl_id == 1:
            A[2, 0] = DW["dhwl"]
            A[2, 1] = DW["hhwl"]
            A[2, 2] = (1 + DW["a3wg1"]) * DW["hhwl"] - DW["a3wg1"] * dwl / 100
            A[2, 3] = 0
            b[2] = dwl
        if wl_id == 2:
            A[2, 0] = DW["dhwl"]
            A[2, 1] = (1 + DW["hhwg"]) * DW["dhwl"] - DW["hhwg"] * dwl / 100
            A[2, 2] = (1 + DW["a3wg2"]) * DW["dhwl"] - DW["a3wg2"] * dwl / 100
            A[2, 3] = 0
            b[2] = dwl
        b[0] = dwa
        b[1] = dwh
        b[3] = 100
    else:
        A[0, 0] = 0
        A[0, 1] = 0
        A[0, 2] = -1
        A[0, 3] = 0
        A[1, 0] = 0
        A[1, 1] = DW["hhwg"]
        A[1, 2] = -1
        A[1, 3] = 0
        A[3, 0] = 1
        A[3, 1] = 1
        A[3, 2] = 1
        A[3, 3] = 1
        if wl_id == 0:
            A[2, 0] = DW["dhwl"]
            A[2, 1] = DW["hhwl"]
            A[2, 2] = 1
            A[2, 3] = 0
            b[2] = dwl
        if wl_id == 1:
            A[2, 0] = DW["dhwl"]
            A[2, 1] = DW["hhwl"]
            A[2, 2] = 0.01 * dwl
            A[2, 3] = 0
            b[2] = dwl
        if wl_id == 2:
            A[2, 0] = DW["dhwl"]
            A[2, 1] = (1 + DW["hhwg"]) * DW["dhwl"] - DW["hhwg"] * dwl / 100
            A[2, 2] = dwl / 100
            A[2, 3] = 0
            b[2] = dwl
        b[0] = dwa
        b[1] = dwh
        b[3] = 100
    return A, b


def tga_linear_old(dwo, dwa, dwh):
    """
    Function to calculate components of system of linear equations giving stucco phase content from three weight loss
    measurements.
    :param dwo: ORG sample weight loss
    :param dwa: HUM sample weight loss
    :param dwh: HYD sample weight loss
    :return:
    A: 4x4 matrix containing coefficients of system of linear equations giving X=[DH, HH, AIII, FM, IN]
    b: array of length 4 giving the constants in AX - b = 0 equation.
    """
    A = np.zeros((4, 4))
    b = np.zeros(4)
    if dwa >= dwo:
        A[0, 0] = 1
        A[0, 1] = 1
        A[0, 2] = 1
        A[0, 3] = 1
        A[1, 0] = DW["dhwl"]
        A[1, 1] = DW["hhwl"]
        A[1, 2] = 0
        A[1, 3] = 0
        A[2, 0] = 0
        A[2, 1] = DW["hhwg"] - 0.01 * DW["hhwg"] / DW["dhwl"] * dwh
        A[2, 2] = DW["a3wg2"] - 0.01 * DW["a3wg2"] / DW["dhwl"] * dwh
        A[2, 3] = -1
        A[3, 0] = DW["dhwl"] / DW["hhwl"] - 1
        A[3, 1] = 0
        A[3, 2] = DW["a3wg1"] - 0.01 * DW["a3wg1"] / DW["hhwl"] * dwa
        A[3, 3] = -1
        b[0] = 100
        b[1] = dwo
        b[2] = 1 / DW["dhwl"] * dwh - 100
        b[3] = 1 / DW["hhwl"] * dwa - 100
    else:
        A[0, 0] = 1
        A[0, 1] = 1
        A[0, 2] = 1
        A[0, 3] = 1
        A[1, 0] = DW["dhwl"]
        A[1, 1] = DW["hhwl"]
        A[1, 2] = 1
        A[1, 3] = 0
        A[2, 0] = 0
        A[2, 1] = DW["hhwg"] - 0.01 * DW["hhwg"] / DW["dhwl"] * dwh
        A[2, 2] = 0.01 * dwh / DW["dhwl"] - 1
        A[2, 3] = -1
        A[3, 0] = DW["dhwl"] / DW["hhwl"] - 1
        A[3, 1] = 0
        A[3, 2] = 0.01 * dwa / DW["hhwl"] - 1
        A[3, 3] = -1
        b[0] = 100
        b[1] = dwo
        b[2] = 1 / DW["dhwl"] * dwh - 100
        b[3] = 1 / DW["hhwl"] * dwa - 100
    return A, b


def tga_linear(dwo, dwa, dwh):
    """
    Function to calculate components of system of linear equations giving stucco phase content from three weight loss
    measurements.
    :param dwo: ORG sample weight loss
    :param dwa: HUM sample weight loss
    :param dwh: HYD sample weight loss
    :return:
    A: 4x4 matrix containing coefficients of system of linear equations giving X=[DH, HH, AIII, FM, IN]
    b: array of length 4 giving the constants in AX - b = 0 equation.
    """
    A = np.zeros((4, 4))
    b = np.zeros(4)
    if dwa >= dwo:
        A[0, 0] = DW["dhwl"]
        A[0, 1] = DW["hhwl"]
        A[0, 2] = 0
        A[0, 3] = 0
        A[1, 0] = DW["dhwl"]
        A[1, 1] = DW["hhwl"]
        A[1, 2] = DW["hhwl"]+DW["hhwl"]*DW["a3wg1"] - 0.01*DW["a3wg1"]*dwa
        A[1, 3] = 0
        A[2, 0] = DW["dhwl"]
        A[2, 1] = DW["dhwl"]+DW["dhwl"]*DW["hhwg"]- 0.01*DW["hhwg"]*dwh
        A[2, 2] = DW["dhwl"]+DW["dhwl"]*DW["a3wg2"]- 0.01*DW["a3wg2"]*dwh
        A[2, 3] = 0
        A[3, 0] = 1
        A[3, 1] = 1
        A[3, 2] = 1
        A[3, 3] = 1
        b[0] = dwo
        b[1] = dwa
        b[2] = dwh
        b[3] = 100
    else:
        A[0, 0] = DW["dhwl"]
        A[0, 1] = DW["hhwl"]
        A[0, 2] = 1
        A[0, 3] = 0
        A[1, 0] = DW["dhwl"]
        A[1, 1] = DW["hhwl"]
        A[1, 2] = 0.01*DW["a3wg1"]*dwa
        A[1, 3] = 0
        A[2, 0] = DW["dhwl"]
        A[2, 1] = DW["dhwl"]+DW["dhwl"]*DW["hhwg"]- 0.01*DW["hhwg"]*dwh
        A[2, 2] = 0.01*DW["a3wg2"]*dwh
        A[2, 3] = 0
        A[3, 0] = 1
        A[3, 1] = 1
        A[3, 2] = 1
        A[3, 3] = 1
        b[0] = dwo
        b[1] = dwa
        b[2] = dwh
        b[3] = 100
    return A, b


def load_excel(filename, test_type):
    """
    Function to read measurements from Excel files
    :param
    filename: the name of the Excel file
    type: "GPA" or "TGA"

    :return:
    labels: sample labels read from the first sheet
    org: measurements for original sample
    hum: measurements for HUM sample
    hyd: measurements for HYF samples

    Note: The Excel file should have a specific format. Use the provided templates or make the excel files according
    to following format. The Excel file should contain 4 sheets as follows:
    Sheet 1: Labels
    Starting from A1, column A should contain the labels of stucco samples analyzed. If you analyzed N samples, A1 to
    AN should have values.
    Sheet 2: Original samples (ORG)
    Column A would be the same as Sheet1. For TGA and COMPUTRAC each row of this sheet should contain weight loss
    measurements (%) for original samples corresponding to labels in sheet 1. It is highly recommended that three
    measurements are reported for all samples although the code can handle lower number of measurements. For GPA
    analysis, this sheet will be ignored.
    Sheet 3: Humidified and dried samples (HUM)
    Column A would be the same as Sheet1. For TGA and COMPUTRAC each row of this sheet should contain weight loss
    measurements (%) for AIII samples (original samples kept in 75% humidity at 45 C over night and dried at 45 C for
    2 hours. For the digital scale measurements, each row contains weight gain percentage of AIII samples.
    Sheet 4: Fully hydrated and dried samples (HYD)
    Column A would be the same as Sheet1. For TGA and COMPUTRAC each row of this sheet should contain weight loss
    measurements (%) for fully hydrated samples (original fully hydrated and dried at 45 C for 24  hours). For the
    digital scale measurements, each row contains weight gain percentage of hydrated samples.

    """

    wb = xlrd.open_workbook(filename)
    sheet = wb.sheet_by_index(0)
    labels = []
    sample_count = sheet.nrows
    for i in range(1, sample_count, 1):
        labels.append(sheet.cell_value(i, 0))

    if test_type == "TGA":
        sheet = wb.sheet_by_index(1)
        org = [[] for i in range(len(labels))]
        for i in range(1, sample_count, 1):
            for j in range(1, min(4, sheet.ncols)):
                org[i - 1].append(sheet.cell_value(i, j))
    else:
        org = None

    sheet = wb.sheet_by_index(2)
    hum = [[] for i in range(len(labels))]
    for i in range(1, sample_count, 1):
        for j in range(1, min(4, sheet.ncols)):
            hum[i - 1].append(sheet.cell_value(i, j))

    sheet = wb.sheet_by_index(3)
    hyd = [[] for i in range(len(labels))]
    for i in range(1, sample_count, 1):
        for j in range(1, min(4, sheet.ncols)):
            hyd[i - 1].append(sheet.cell_value(i, j))

    return labels, org, hum, hyd


def tga_solve(tga_measurements):
    """
    Function to calculate phase contents of stucco from weight loss measurements.

    :param tga_measurements: tuple including following lists:
    tga_labels: A list containing samples labels
    tga_org: A list of lists containing weight loss measurements of ORG samples for each label
    tga_hum: A list of lists containing weight loss measurements of HUM samples for each label
    tga_hyd: A list of lists containing weight loss measurements of HYD samples for each label
    :return: phases: A dictionary. Keys: labels. Values: [mean, std]
    means: [DH.mean, HH.mean, AIII.mean, FM.mean, IN.mean] (mean value of phases)
    """

    tga_labels, tga_org, tga_hum, tga_hyd = tga_measurements

    tga_pa = [[] for i in range(len(tga_labels))]
    lbond = np.zeros(4)
    ubond = lbond + 100

    for i1 in range(len(tga_labels)):
        if tga_labels[i1]:
            for i2 in range(len(tga_org[i1])):
                if isdigit(str(tga_org[i1][i2])):
                    for i3 in range(len(tga_hum[i1])):
                        if isdigit(str(tga_hum[i1][i3])):
                            for i4 in range(len(tga_hyd[i1])):
                                if isdigit(str(tga_hyd[i1][i4])) and i2 == i3 and i2 == i4:
                                    A, b = tga_linear(tga_org[i1][i2], tga_hum[i1][i3], tga_hyd[i1][i4])
                                    # p = np.linalg.solve(A, b)
                                    lsd = lsq_linear(A, b, bounds=(lbond, ubond), verbose=0)
                                    p = list(lsd.x)
                                    if tga_hum[i1][i3] >= tga_org[i1][i2]:
                                        p.insert(3, 0)
                                    else:
                                        p.insert(2, 0)
                                    tga_pa[i1].append(p)
    tga_pa = [np.array(x) for x in tga_pa]

    phases = {}
    for i in range(len(tga_labels)):
        phases[tga_labels[i]] = dict(zip(["mean", "std"], [np.mean(tga_pa[i], axis=0).tolist(),
                                                           np.std(tga_pa[i], axis=0).tolist()]))
        # phases[tga_labels[i]]["std"] = np.std(tga_pa[0], axis=0)

    return phases


def gpa_solve(tga_measurements, gpa_measurements):
    """
    Function to calculate phase contents of stucco from one weight loss and two weight gain measurements.

    :param tga_measurements: tuple including following lists:
    tga_labels: A list containing samples labels
    tga_org: A list of lists containing weight loss measurements of ORG samples for each label
    tga_hum: A list of lists containing weight loss measurements of HUM samples for each label
    tga_hyd: A list of lists containing weight loss measurements of HYD samples for each label
    :param gpa_measurements: tuple including following lists:
    gpa_labels: A list of samples labels for weight gain measurements
    gpa_hum: A list of lists containing weight gain measurements of HUM samples for each label
    gpa_hyd: A list of lists containing weight gain measurements of HYD samples for each label
    :return: phases: A dictionary. Keys: labels. Values: [mean, std]
    means: [DH.mean, HH.mean, AIII.mean, FM.mean, IN.mean] (mean value of phases)

    None: the weight loss measurements is used to calculate DH content. It won't affect HH and AIII/FM.
    """

    gpa_labels, gpa_org, gpa_hum, gpa_hyd = gpa_measurements
    tga_labels, tga_org, tga_hum, tga_hyd = tga_measurements

    gpa_pa_org = [[] for i in range(len(gpa_labels))]
    gpa_pa_hum = [[] for i in range(len(gpa_labels))]
    gpa_pa_hyd = [[] for i in range(len(gpa_labels))]

    missmatches = []
    lbond = np.zeros(4)
    ubond = lbond + 100

    for tga_wl_id in range(3):
        for i1 in range(len(gpa_labels)):
            if gpa_labels[i1] and i1 < len(tga_labels):
                if gpa_labels[i1] != tga_labels[i1]:
                    missmatches.append([gpa_labels[i1]])
                    continue
                else:
                    tga_wl = tga_measurements[tga_wl_id+1]
                    for i2 in range(len(tga_wl[i1])):
                        if isdigit(str(tga_wl[i1][i2])):
                            for i3 in range(len(gpa_hum[i1])):
                                if isdigit(str(gpa_hum[i1][i3])):
                                    for i4 in range(len(gpa_hyd[i1])):
                                        if isdigit(str(gpa_hyd[i1][i4])) and i2 == i3 and i2 == i4:
                                            A, b = gpa_linear(tga_wl_id, tga_wl[i1][i2], gpa_hum[i1][i3], gpa_hyd[i1][i4])
                                            #p = list(np.linalg.solve(A, b))
                                            lsd = lsq_linear(A, b, bounds=(lbond, ubond), lsmr_tol='auto', verbose=0)
                                            p = list(lsd.x)
                                            if gpa_hum[i1][i3] >= 0:
                                                p.insert(3, 0)
                                            else:
                                                p.insert(2, 0)
                                            if tga_wl_id == 0:
                                                gpa_pa_org[i1].append(p)
                                            elif tga_wl_id == 1:
                                                gpa_pa_hum[i1].append(p)
                                            else:
                                                gpa_pa_hyd[i1].append(p)

    gpa_pa_org = [np.array(x) for x in gpa_pa_org]
    gpa_pa_hum = [np.array(x) for x in gpa_pa_hum]
    gpa_pa_hyd = [np.array(x) for x in gpa_pa_hyd]
    gpa_pa = np.concatenate((gpa_pa_org, gpa_pa_hum, gpa_pa_hyd))
    phases_org = {}
    phases_hum = {}
    phases_hyd = {}
    phases={}
    for i in range(len(tga_labels)):
        phases_org[gpa_labels[i]] = dict(zip(["mean", "std"], [np.mean(gpa_pa_org[i], axis=0).tolist(),
                                                           np.std(gpa_pa_org[i], axis=0).tolist()]))
        phases_hum[gpa_labels[i]] = dict(zip(["mean", "std"], [np.mean(gpa_pa_hum[i], axis=0).tolist(),
                                                               np.std(gpa_pa_hum[i], axis=0).tolist()]))
        phases_hyd[gpa_labels[i]] = dict(zip(["mean", "std"], [np.mean(gpa_pa_hyd[i], axis=0).tolist(),
                                                               np.std(gpa_pa_hyd[i], axis=0).tolist()]))
        phases[gpa_labels[i]] = dict(zip(["mean", "std"], [np.mean(gpa_pa[i], axis=0).tolist(),
                                                               np.std(gpa_pa[i], axis=0).tolist()]))
    return phases, phases_org, phases_hum, phases_hyd


def save_excel(address, phases):
    """

    :param address: address at which the output excel file should be saved
    :param phases: a phases dictionary:
    labels: name of the analysis (like "TGA" or "GPA"). This will be the name of the sheet.
    values: phase calculations outputs of tga_solve or gpa_solve functions. These are dictionaries whose labels are
    sample labels and values are mean and std of each phase content.
    :return: None
    """

    color = ['light_blue', 'green', 'light_orange', 'light_blue', 'green']
    workbook = xlwt.Workbook()
    style = xlwt.XFStyle()
    style.alignment.wrap = 1
    filename = Path(address) / 'Phase_Contents.xls'

    for analysis in phases:
        if not isListEmpty(phases[analysis]):
            sheet = workbook.add_sheet(analysis)
            pattern = "pattern: pattern solid, fore_color " + color[
                0] + "; font: color black, bold on; align: horiz center"
            sheet.write_merge(0, 0, 1, 2, 'DH', xlwt.easyxf(pattern))
            pattern = "pattern: pattern solid, fore_color " + color[
                1] + "; font: color black, bold on; align: horiz center"
            sheet.write_merge(0, 0, 3, 4, 'HH', xlwt.easyxf(pattern))
            pattern = "pattern: pattern solid, fore_color " + color[
                2] + "; font: color black, bold on; align: horiz center"
            sheet.write_merge(0, 0, 5, 6, 'AIII', xlwt.easyxf(pattern))
            pattern = "pattern: pattern solid, fore_color " + color[
                3] + "; font: color black, bold on; align: horiz center"
            sheet.write_merge(0, 0, 7, 8, 'FM', xlwt.easyxf(pattern))
            pattern = "pattern: pattern solid, fore_color " + color[
                4] + "; font: color black, bold on; align: horiz center"
            sheet.write_merge(0, 0, 9, 10, 'IN', xlwt.easyxf(pattern))
            for ii in range(5):
                pattern = "pattern: pattern solid, fore_color " + color[
                    ii] + "; font: color white; font: bold 1; align: horiz center"
                sheet.write(1, ii * 2 + 1, 'Mean', xlwt.easyxf(pattern))
                sheet.write(1, ii * 2 + 2, 'STDEV', xlwt.easyxf(pattern))
            length = 0
            i1 = 0
            p = phases[analysis]
            for label in p:
                if label:
                    if len(str(label)) > length:
                        length = len(str(label))
                    sheet.write(2 + i1, 0, str(label))
                for i2 in range(5):
                    sheet.write(2 + i1, 1 + i2*2, p[label]['mean'][i2])
                    sheet.write(2 + i1, 2 + i2 * 2, p[label]['std'][i2])
                i1 += 1
            first_col = sheet.col(0)
            first_col.width = 256 * length
    workbook.save(filename)



