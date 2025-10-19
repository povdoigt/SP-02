import xlwings
import numpy as np
from matplotlib import pyplot as plt


TRAJECTO_RAMPE_ANGLE = 'C20'
TRAJECTO_ALT_Z = 'I27'
TRAJECTO_PORTEE_X = 'J29'

ab_wb = xlwings.Book('alpha_beta.xlsx')
ab_trajecto = ab_wb.sheets['Trajecto']
ab_Calculs = ab_wb.sheets['Calculs']

b_wb = xlwings.Book('beta.xlsx')
b_trajecto = b_wb.sheets['Trajecto']

angles = np.linspace(45, 85, 10)

ab_alt_z_values = []
ab_portee_x_values = []
ab_vit_max_values =[]
ab_acc_max_values =[]

b_alt_z_values = []
b_portee_x_values = []
b_vit_max_values = []
b_acc_max_values = []

ang_sepa_values = []

for angle in angles:
    ab_trajecto.range(TRAJECTO_RAMPE_ANGLE).value = angle
    alt_z = ab_trajecto.range(TRAJECTO_ALT_Z).value
    portee_x = ab_trajecto.range(TRAJECTO_PORTEE_X).value
    vit_max = ab_trajecto.range('K25').value / 340.3
    acc_max = ab_trajecto.range('L25').value / 9.81
    
    ab_alt_z_values.append(alt_z)
    ab_portee_x_values.append(portee_x)
    ab_vit_max_values.append(vit_max)
    ab_acc_max_values.append(acc_max)

    pos_z = ab_Calculs.range('K221').value
    pos_x = ab_Calculs.range('J221').value
    vit_xz = ab_Calculs.range('I221').value
    ang_sepa = ab_Calculs.range('N221').value

    ang_sepa_values.append(ang_sepa)

    b_trajecto.range('I42').value = pos_z
    b_trajecto.range('J42').value = pos_x
    b_trajecto.range('K42').value = vit_xz
    b_trajecto.range(TRAJECTO_RAMPE_ANGLE).value = ang_sepa

    alt_z = b_trajecto.range(TRAJECTO_ALT_Z).value
    portee_x = b_trajecto.range(TRAJECTO_PORTEE_X).value
    vit_max = b_trajecto.range('K25').value / 340.3
    acc_max = b_trajecto.range('L25').value / 9.81

    b_alt_z_values.append(alt_z)
    b_portee_x_values.append(portee_x)
    b_vit_max_values.append(vit_max)
    b_acc_max_values.append(acc_max)

    print(f"Angle: {angle}°")

angle = 80

ab_trajecto.range(TRAJECTO_RAMPE_ANGLE).value = angle
alt_z = ab_trajecto.range(TRAJECTO_ALT_Z).value
portee_x = ab_trajecto.range(TRAJECTO_PORTEE_X).value
vit_max = ab_trajecto.range('K25').value / 340.3
acc_max = ab_trajecto.range('L25').value / 9.81

pos_z = ab_Calculs.range('K231').value
pos_x = ab_Calculs.range('J231').value
vit_xz = ab_Calculs.range('I231').value
ang_sepa = ab_Calculs.range('N231').value

b_trajecto.range('I42').value = pos_z
b_trajecto.range('J42').value = pos_x
b_trajecto.range('K42').value = vit_xz
b_trajecto.range(TRAJECTO_RAMPE_ANGLE).value = ang_sepa

alt_z = b_trajecto.range(TRAJECTO_ALT_Z).value
portee_x = b_trajecto.range(TRAJECTO_PORTEE_X).value
vit_max = b_trajecto.range('K25').value / 340.3
acc_max = b_trajecto.range('L25').value / 9.81

plt.figure(figsize=(10, 10))

plt.subplot(2, 2, 1)
plt.plot(angles, ab_alt_z_values, label='Alpha Beta')
plt.plot(angles, b_alt_z_values, label='Beta', linestyle='--')
plt.title('Altitude Z vs Angle')
plt.xlabel('Angle (degrees)')
plt.ylabel('Altitude Z (m)')
plt.grid()
plt.legend()

plt.subplot(2, 2, 2)
plt.plot(angles, ab_portee_x_values, label='Alpha Beta')
plt.plot(angles, b_portee_x_values, label='Beta', linestyle='--')
plt.title('Portée X vs Angle')
plt.xlabel('Angle (degrees)')
plt.ylabel('Portée X (m)')
plt.grid()
plt.legend()

plt.subplot(2, 2, 3)
plt.plot(angles, ab_vit_max_values, label='Alpha Beta')
plt.plot(angles, b_vit_max_values, label='Beta', linestyle='--')
plt.title('Vitesse Max vs Angle')
plt.xlabel('Angle (degrees)')
plt.ylabel('Vitesse Max (Mach 0 m)')
plt.grid()
plt.legend()

plt.subplot(2, 2, 4)
plt.plot(angles, ab_acc_max_values, label='Alpha Beta')
plt.plot(angles, b_acc_max_values, label='Beta', linestyle='--')
plt.title('Accélération Max vs Angle')
plt.xlabel('Angle (degrees)')
plt.ylabel('Accélération Max (g)')
plt.grid()
plt.legend()

plt.tight_layout()
plt.show()

# print(ab_Calculs.range('B430').value)

