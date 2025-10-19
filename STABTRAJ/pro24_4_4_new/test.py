import xlwings
import numpy as np
from matplotlib import pyplot as plt
import time


TRAJECTO_RAMPE_ANGLE = 'C20'
TRAJECTO_ALT_Z = 'I27'
TRAJECTO_PORTEE_X = 'J29'

ab_wb = xlwings.Book('alpha_beta.xlsx')
ab_trajecto = ab_wb.sheets['Trajecto']
ab_Calculs = ab_wb.sheets['Calculs']

b_wb = xlwings.Book('beta.xlsx')
b_trajecto = b_wb.sheets['Trajecto']


alt_z = b_trajecto.range(TRAJECTO_ALT_Z).value

print(f"Initial Altitude Z: {alt_z:.2f} m")