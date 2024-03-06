import xlrd
import numpy as np
import pandas as pd
import seaborn as sns
from fig_settings import *

# read different EoL teartment strategies data
def get_data_from_recy_new(path="input_data/Wind_data.xls"):
    wb = xlrd.open_workbook(path)
    # load existing data
    sheet = wb.sheet_by_name('recy_rate_new')
    
    table = {}
    
    for i in range(41):
        
        if 'EoL' in sheet.cell_value(i, 0):
            start_row = i
            t_type = sheet.cell_value(i, 0)
            
            # create t type table
            table[t_type + '_onshore'] = {}
            table[t_type + '_offshore'] = {}
            
            for b_i in range(i + 2, i + 12):
                material = sheet.cell_value(b_i, 1)
                onshore_recy = [sheet.cell_value(b_i, j) for j in range(2, 7)]
                offshore_recy = [sheet.cell_value(b_i, j) for j in range(9, 14)]
                
                # write to table
                table[t_type + '_onshore'][material] = onshore_recy
                table[t_type + '_offshore'][material] = offshore_recy

    proc_methods = [sheet.cell_value(2, j).replace('on_', '') for j in range(2, 7)]
    
    # print table to verify
    print(table)
    
    return table, proc_methods
    
#  read the onshore wind turbine material consumption data
def load_onshore_dict(path="input_data/Wind_data.xls"):
    '''
    return a dictionary of onshore data
    {Nacc: ..., Tower: ..., Blade: ..., Hub: ...}
    '''
    
    wb = xlrd.open_workbook(path)
    sheet = wb.sheet_by_name('on_material')
    onshore_dict = {}
    for i in range(3, 14):
        m = sheet.cell_value(i, 0)
        if len(m) == 0:
            continue
        for j in range(1, 11):
            t = sheet.cell_value(1, j)
            if t not in onshore_dict:
                onshore_dict[t] = {}
            if m not in onshore_dict[t]:
                onshore_dict[t][m] = {}
            onshore_dict[t][m] = sheet.cell_value(i, j) if sheet.cell_value(i, j) != '' else 0
            if i >= 12:
                onshore_dict[t][m] = onshore_dict[t][m] / 10**6
    return onshore_dict

# read the offshore wind turbine material consumption data
def load_offshore_dict(path="input_data/Wind_data.xls"):
    '''
    return a dictionary of onshore data
    {Nacc: ..., Tower: ..., Blade: ..., Hub: ...}
    '''
    
    wb = xlrd.open_workbook(path)
    sheet = wb.sheet_by_name('off_material')
    onshore_dict = {}
    for i in range(3, 14):
        m = sheet.cell_value(i, 0)
        if len(m) == 0:
            continue
        for j in range(1, 11):
            t = sheet.cell_value(1, j)
            if t not in onshore_dict:
                onshore_dict[t] = {}
            if m not in onshore_dict[t]:
                onshore_dict[t][m] = {}
            onshore_dict[t][m] = sheet.cell_value(i, j) if sheet.cell_value(i, j) != '' else 0
            if i >= 12:
                onshore_dict[t][m] = onshore_dict[t][m] / 10**6
    return onshore_dict

# read the histortcal onshore wind turbine information 
def load_original_data(path="input_data/Wind_data.xls"):
    '''
    :param path: path to the excel file
    :return: list of d, h, nacelle, tower, time
    '''
    wb = xlrd.open_workbook(path)
    # load existing data
    sheet = wb.sheet_by_name('historical_info')
    n_rows = 6698
    c_list = [sheet.cell_value(i, 5) for i in range(1, n_rows)]
    d_list = [sheet.cell_value(i, 6) for i in range(1, n_rows)]
    h_list = [sheet.cell_value(i, 7) for i in range(1, n_rows)]
    nacl_list = [sheet.cell_value(i, 11) for i in range(1, n_rows)]
    tower_list = [sheet.cell_value(i, 12) for i in range(1, n_rows)]
    time_list = [sheet.cell_value(i, 10) for i in range(1, n_rows)]
    
    for i in range(len(nacl_list)):
        if 'DFIG' in nacl_list[i] or 'SCIG' in nacl_list[i]:
            nacl_list[i] = 'DFIG/SCIG'
    for i in range(len(c_list)):
        if '/' in str(c_list[i]):
            # get average
            c_list[i] = np.mean([float(x) for x in c_list[i].split('/')])
        if '-' in str(c_list[i]):
            # get average
            c_list[i] = np.mean([float(x) for x in c_list[i].split('-')])
    
    return c_list, d_list, h_list, nacl_list, tower_list, time_list
    
# calculate the mass of each component of onshore wind turbine
def calculate_total_mass(d, h, sec):
    if sec == 'Nacelle':
        mass = 0.0091 * d ** (2.0456)
    elif sec == 'Tower':
        mass = 0.0176*(d ** 2 * h) ** 0.6839
    elif sec == 'Rotor':
        mass = 0.0035 * d ** 2.1412
    elif sec == 'Foundation':
        mass = 3.5 * (0.0091 * d ** (2.0456) \
                        + 0.0176*(d ** 2 * h) ** 0.6839 \
                        + 0.0035 * d ** 2.1412) # 3.5 is the ratio of foundation mass to total mass
    else:
        raise ValueError('Unknown section: {}'.format(sec))
    return mass

# calculate the mass of each component of offshore wind turbine
def calculate_offshore_mass(d, h, t, sec):
    if sec == 'Nacelle':
        mass = 0.0091 * d ** (2.0456)
    elif sec == 'Tower':
        mass = 0.0176*(d ** 2 * h) ** 0.6839
    elif sec == 'Rotor':
        mass = 0.0035 * d ** 2.1412
    elif sec == 'Foundation':
        scale = 2.2 if int(t) <= 2035 else 2.8
        mass = scale * (0.0091 * d ** (2.0456) \
                        + 0.0176*(d ** 2 * h) ** 0.6839 \
                        + 0.0035 * d ** 2.1412) # 2.2 or 2.8 is the ratio of foundation mass to total mass
    else:
        raise ValueError('Unknown section: {}'.format(sec))
    return mass

# calculate the historical mass of each material for onshore wind turbine
def calculate_material_mass(c, d, h, nacl, tower, oh_dict, nacl_rep, rotor_sum):
    

    f_k = 'flat'
    mass_dict = {m : 0 for m in oh_dict['DFIG/SCIG']}
    for m in mass_dict:
        if m not in ['Nd', 'Dy']:
            nacl_m = calculate_total_mass(d, h, 'Nacelle') * oh_dict[nacl][m] * (nacl_rep/100+1)
            tower_m = calculate_total_mass(d, h, 'Tower') * oh_dict[tower][m]
            rotor_m = calculate_total_mass(d, h, 'Rotor') * oh_dict['/'][m] * (rotor_sum/100+1)
            found_m = calculate_total_mass(d, h, 'Foundation') * oh_dict[f_k][m]
            mass_dict[m] = nacl_m + tower_m + rotor_m + found_m
        else:
            mass_dict[m] = c * oh_dict[nacl][m] \
                           + c * oh_dict[tower][m] \
                           + c * oh_dict['/'][m] \
                           + c * oh_dict[f_k][m]
    return mass_dict

# calculate the future mass of each material under different technology development scenarios for onshore wind turbine
def calculate_future_material_onshore_mass(n, c, d, h, oh_dict, tp=1):
    
    excel_path = "input_data/Wind_data.xls"
    df = pd.read_excel(excel_path, sheet_name='tech_dev', header=1)
    nacl_list = df['nacelle_onshore_summary'].dropna().tolist()
    nacl_vals = df['Unnamed: 3'].dropna().tolist()

    nacl_type = {}
    for i in range(len(nacl_list)):
        if nacl_list[i] not in nacl_type:
            nacl_type[nacl_list[i]] = []
        nacl_type[nacl_list[i]].append(nacl_vals[i])

    nacl_type = {k: v[tp] for k, v in nacl_type.items()}

    nacl_rep = df['nacelle_replacement_summary'].dropna().tolist()[tp]
    rotor_sum = df['rotor_O _M_summary'].dropna().tolist()[tp]
    tower_type = {
        'Hybrid': df['Hybrid'].dropna().values[0],
        'Steel': df['Steel'].dropna().values[0],
    }
    
    mass_dict = {m : 0 for m in oh_dict['DFIG/SCIG']}
    for m in mass_dict:
        if m not in ['Nd', 'Dy']:
            nacl_m, tower_m = 0, 0
            for t in nacl_type:
                nacl_m += calculate_total_mass(d, h, 'Nacelle') * oh_dict[t][m] * nacl_type[t] * (nacl_rep/100+1)
            for t in tower_type:
                tower_m += calculate_total_mass(d, h, 'Tower') * oh_dict[t][m] * tower_type[t]

            rotor_m = calculate_total_mass(d, h, 'Rotor') * oh_dict['/'][m] * (rotor_sum/100+1)
            found_m = calculate_total_mass(d, h, 'Foundation') * oh_dict['flat'][m]
            mass_dict[m] = nacl_m + tower_m + rotor_m + found_m
        else:
            for t in nacl_type:
                mass_dict[m] += c * oh_dict[t][m] * nacl_type[t] * (nacl_rep/100+1)
            for t in tower_type:
                mass_dict[m] += c * oh_dict[t][m] * tower_type[t]
            mass_dict[m] += c * oh_dict['/'][m] * (rotor_sum/100+1)\
                           + c * oh_dict['flat'][m]
            
    return mass_dict

# calculate the future mass of each material under different technology development scenarios for offshore wind turbine
def calculate_future_material_offshore_mass(n, c, d, h, year, oh_dict,tp=1):
    
    excel_path = "input_data/Wind_data.xls"
    df = pd.read_excel(excel_path, sheet_name='tech_dev', header=1)
    nacl_list = df['nacelle_onshore_summary'].dropna().tolist()
    nacl_vals = df['Unnamed: 5'].dropna().tolist()

    nacl_type = {}
    for i in range(len(nacl_list)):
        if nacl_list[i] not in nacl_type:
            nacl_type[nacl_list[i]] = []
        nacl_type[nacl_list[i]].append(nacl_vals[i])

    nacl_type = {k: v[tp] for k, v in nacl_type.items()}

    nacl_rep = df['nacelle_replacement_summary'].dropna().tolist()[tp]
    rotor_sum = df['rotor_O _M_summary'].dropna().tolist()[tp]
    tower_type = {
        'Hybrid': df['Hybrid'].dropna().values[0],
        'Steel': df['Steel'].dropna().values[0],
    }
    
    f_k = 'flat' if 'flat' in oh_dict else 'Monopile'
    
    mass_dict = {m : 0 for m in oh_dict['DFIG/SCIG']}
    for m in mass_dict:
        if m not in ['Nd', 'Dy']:
            nacl_m, tower_m = 0, 0
            for t in nacl_type:
                nacl_m += calculate_offshore_mass(d, h, year, 'Nacelle') * oh_dict[t][m] * nacl_type[t] * (nacl_rep/100+1)
            for t in tower_type:
                tower_m += calculate_offshore_mass(d, h, year, 'Tower') * oh_dict[t][m] * tower_type[t]

            rotor_m = calculate_offshore_mass(d, h, year, 'Rotor') * oh_dict['/'][m] * (rotor_sum/100+1)
            found_m = calculate_offshore_mass(d, h, year, 'Foundation') * oh_dict[f_k][m]
            mass_dict[m] = nacl_m + tower_m + rotor_m + found_m
        else:
            for t in nacl_type:
                mass_dict[m] += c * oh_dict[t][m] * nacl_type[t]*(nacl_rep/100+1)
            for t in tower_type:
                mass_dict[m] += c * oh_dict[t][m] * tower_type[t]
            mass_dict[m] += c * oh_dict['/'][m] *(rotor_sum/100+1)\
                           + c * oh_dict[f_k][m]
            
    return mass_dict

# caculate the mass of each component of future onshore wind turbine
def calculate_future_material_mass_onshore_by_year(n_list, c_list, d_list, h_list, oh_dict, time_list, tp):
    mass_by_year = {}
    
    for i in range(len(n_list)):
        year = int(time_list[i])
        n = n_list[i]
        c = c_list[i]
        d = d_list[i]
        h = h_list[i]
        mass_dict = calculate_future_material_onshore_mass(n, c, d, h, oh_dict, tp=tp)
        
        for k in mass_dict:
            mass_dict[k] = mass_dict[k] * n

        mass_by_year[year] = mass_dict
    return mass_by_year

# calculate the mass of each component of future offshore wind turbine
def calculate_future_material_mass_offshore_by_year(n_list, c_list, d_list, h_list, oh_dict, time_list, tp):
    mass_by_year = {}
    
    for i in range(len(n_list)):
        year = int(time_list[i])
        n = n_list[i]
        c = c_list[i]
        d = d_list[i]
        h = h_list[i]
        mass_dict = calculate_future_material_offshore_mass(n, c, d, h, year, oh_dict, tp=tp)
        
        for k in mass_dict:
            mass_dict[k] = mass_dict[k] * n

        mass_by_year[year] = mass_dict
    return mass_by_year

# calculate the mass of each component of historical onshore wind turbine
def calculate_material_mass_by_year(c_list, d_list, h_list, nacl_list, tower_list, time_list, oh_dict):
    '''
    :param d_list: list of diameters
    :param h_list: list of hub heights
    :param nacl_list: list of nacelle types
    :param tower_list: list of tower types
    :param time_list: list of time
    :param oh_dict: dictionary of onshore data
    :return: list of materials
    '''
    mass_by_year = {}

    excel_path = "input_data/Wind_data.xls"
    df = pd.read_excel(excel_path, sheet_name='tech_dev', header=1)

    nacl_rep = df['nacelle_replacement_summary'].dropna().tolist()[0]
    rotor_sum = df['rotor_O _M_summary'].dropna().tolist()[0]
    
    for i in range(len(time_list)):
        year = int(time_list[i])
        if year not in mass_by_year:
            mass_by_year[year] = {}
            for m in oh_dict['DFIG/SCIG']:
                mass_by_year[year][m] = 0
        c = c_list[i]
        d = d_list[i]
        h = h_list[i]
        mass_dict = calculate_material_mass(c, d, h, nacl_list[i], tower_list[i], oh_dict, nacl_rep, rotor_sum)
        debug = 0
        for m in mass_dict:
            mass_by_year[year][m] += mass_dict[m]
    # 0 for missing years
    for i in [1994, 1996]:
        if i not in mass_by_year:
            mass_by_year[i] = {}
            for m in oh_dict['DFIG/SCIG']:
                mass_by_year[i][m] = 0
    return mass_by_year


def bar_plot_mass_by_year(mass_by_year, name, w_scale=1):
    materials = list(mass_by_year.columns)
    mass_by_year['Year'] = mass_by_year.index
    
    fig, axes = plt.subplots(figsize=(w_scale*FIG_WIDTH, 8*FIG_HEIGHT), nrows=10)
    
    for m in materials:
        ax = axes[materials.index(m)]
        # bar plot
        sns.barplot(x='Year', y=m, data=mass_by_year, ax=ax, color=COLORS[m])
        #ax.set_xlabel('Year')
        # rotate xticks
        #ax.set_xticks(ax.get_xticks())
        #ax.set_xticklabels(ax.get_xticklabels(), rotation=45)
        # fit the trend line
        # remove right and top spines
        ax.spines['right'].set_visible(False)
        ax.spines['top'].set_visible(False)
        
    # save to pdf
    fig.tight_layout()
    plt.savefig(name)
    
    
def plot_mass_by_year(mass_by_year, name, w_scale=1):
    materials = list(mass_by_year.columns)
    mass_by_year['Year'] = mass_by_year.index
    
    fig, ax = plt.subplots(figsize=(1.7*FIG_WIDTH, 1.7*FIG_HEIGHT))
    
    bottom = np.zeros(len(mass_by_year))
    
    mean_v = [-np.mean(mass_by_year[m]) for m in materials]
    # sort materials by mean value
    materials = [m for _, m in sorted(zip(mean_v, materials))]
    
    for m in materials:

        ax.plot(mass_by_year.index, mass_by_year[m] + bottom, label=m, color=COLORS[m])
        
        # fill in color between mass_by_year[m] and bottom
        ax.fill_between(mass_by_year.index, mass_by_year[m] + bottom, bottom, color=COLORS[m])
        
        ax.spines['right'].set_visible(False)
        ax.spines['top'].set_visible(False)
        
        bottom += mass_by_year[m]
    
    ax.legend()
    ax.set_xlabel('Year')
    ax.set_ylabel('Mass [t]')
    
    
    # remove legend frame
    leg = ax.get_legend()
    leg.get_frame().set_linewidth(0.0)
    
    # save to pdf
    fig.tight_layout()
    plt.savefig(name)
