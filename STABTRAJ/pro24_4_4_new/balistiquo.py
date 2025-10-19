import xlwings
import numpy as np
from matplotlib import pyplot as plt
import time


TRAJECTO_RAMPE_ANGLE = 'C20'
TRAJECTO_ALT_Z = 'I27'
TRAJECTO_PORTEE_X = 'J29'


alpha_beta_name = 'Alpha_Beta/alpha_beta_actif.xlsx'
beta_name = 'Beta/beta_actif.xlsx'

ab_wb = xlwings.Book(alpha_beta_name)
ab_trajecto = ab_wb.sheets['Trajecto']
ab_Calculs = ab_wb.sheets['Calculs']

b_wb = xlwings.Book(beta_name)
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

    print(f"Angle: {angle}°")

    ab_trajecto.range(TRAJECTO_RAMPE_ANGLE).value = angle
    alt_z = ab_trajecto.range(TRAJECTO_ALT_Z).value
    portee_x = ab_trajecto.range(TRAJECTO_PORTEE_X).value
    vit_max = ab_trajecto.range('K25').value / 340.3
    acc_max = ab_trajecto.range('L25').value / 9.81

    print(f"Alpha Beta - Altitude Z: {alt_z:.2f} m, Portée X: {portee_x:.2f} m, Vitesse Max: {vit_max:.2f} Mach, Accélération Max: {acc_max:.2f} g")
    
    ab_alt_z_values.append(alt_z)
    ab_portee_x_values.append(portee_x)
    ab_vit_max_values.append(vit_max)
    ab_acc_max_values.append(acc_max)

    pos_z = ab_Calculs.range('K216').value
    pos_x = ab_Calculs.range('J216').value
    vit_xz = ab_Calculs.range('I216').value
    ang_sepa = ab_Calculs.range('N216').value

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

    print(f"Beta       - Altitude Z: {alt_z:.2f} m, Portée X: {portee_x:.2f} m, Vitesse Max: {vit_max:.2f} Mach, Accélération Max: {acc_max:.2f} g\n")

plt.figure(figsize=(10, 10))
plt.suptitle(f'Trajectoire Actif')

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


angle = 80

for i in range(2):
    if i == 0:
        alpha_beta_name = 'Alpha_Beta/alpha_beta_actif.xlsx'
        alpha_name = 'Alpha/alpha_vol_actif.xlsx'
    else:
        alpha_beta_name = 'Alpha_Beta/alpha_beta_passif.xlsx'
        alpha_name = 'Alpha/alpha_vol_passif.xlsx'

    ab_wb = xlwings.Book(alpha_beta_name)
    ab_trajecto = ab_wb.sheets['Trajecto']
    ab_Calculs = ab_wb.sheets['Calculs']

    a_wb = xlwings.Book(alpha_name)
    a_trajecto = a_wb.sheets['Trajecto']

    ab_trajecto.range(TRAJECTO_RAMPE_ANGLE).value = angle

    pos_z = ab_Calculs.range('K216').value
    pos_x = ab_Calculs.range('J216').value
    vit_xz = ab_Calculs.range('I216').value
    ang_sepa = ab_Calculs.range('N216').value

    a_trajecto.range('I42').value = pos_z
    a_trajecto.range('J42').value = pos_x
    a_trajecto.range('K42').value = vit_xz
    a_trajecto.range(TRAJECTO_RAMPE_ANGLE).value = ang_sepa

    if i == 0:
        for j in range(2):
            if j == 0:
                beta_name = 'Beta/beta_actif.xlsx'
            else:
                beta_name = 'Beta/beta_passif_moteur.xlsx'

            b_wb = xlwings.Book(beta_name)
            b_trajecto = b_wb.sheets['Trajecto']

            b_trajecto.range('I42').value = pos_z
            b_trajecto.range('J42').value = pos_x
            b_trajecto.range('K42').value = vit_xz
            b_trajecto.range(TRAJECTO_RAMPE_ANGLE).value = ang_sepa
    else:
        beta_name = 'Beta/beta_passif_vide.xlsx'
        
        b_wb = xlwings.Book(beta_name)
        b_trajecto = b_wb.sheets['Trajecto']

        b_trajecto.range('I42').value = pos_z
        b_trajecto.range('J42').value = pos_x
        b_trajecto.range('K42').value = vit_xz
        b_trajecto.range(TRAJECTO_RAMPE_ANGLE).value = ang_sepa
