import numpy as np
from core import utils
from pathlib import Path

GPA_data_file = \
    Path('G:\My Drive\Gypsum Project\Gypsum Project (2019)\Data and Analyses\Phase Analysis\Feb2020-pure samples\Cleaned\Digital_Scale-C.xlsx')
TGA_data_file = \
    Path('G:\My Drive\Gypsum Project\Gypsum Project (2019)\Data and Analyses\Phase Analysis\Feb2020-pure samples\Cleaned\COM-C.xlsx')
out_address = 'Output'
#TGA_data_file = 'Input/TGA-C.xlsx'
Output_file = 'Output/PA.xlsx'

gpa_measurements = utils.load_excel(GPA_data_file, "GPA")
tga_measurements = utils.load_excel(TGA_data_file, "TGA")

print(tga_measurements)

phases = {}
tga_phases = utils.tga_solve(tga_measurements)
phases.update({'TGA': tga_phases})
gpa_phases, gpa_phases_org, gpa_phases_hum, gpa_phases_hyd = utils.gpa_solve(tga_measurements, gpa_measurements)
phases.update({'GPA': gpa_phases})
phases.update({'GPA_ORG': gpa_phases_org})
phases.update({'GPA_HUM': gpa_phases_hum})
phases.update({'GPA_HYD': gpa_phases_hyd})

utils.save_excel(out_address, phases)





