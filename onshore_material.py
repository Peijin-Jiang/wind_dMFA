import math
from fig_utils import *
from capacity_flow_all import capacity_flow_smooth

# relationship between capacity and height for onshore wind turbine
def get_height(x):
    return 4.1099 * x ** 0.3974

# relationship between capacity and diameter for onshore wind turbine
def get_diameter(x):
    return 2.1464 * x ** 0.4913


def capacity_onshore(tp=1):
    
    excel_path = "input_data/Wind_data.xls"
    df = pd.read_excel(excel_path, sheet_name='on_capacity')
    per_cap_list = df['on_future_capacity_per_turbine '].dropna().tolist()

    future_inflow_on, future_inflow_off, ratio_on, _ = capacity_flow_smooth()

    # assumptions for future onshore average capacity per wind turbine
    capacity_per_wind_20_29 = per_cap_list[0]
    capacity_per_wind_30_39 = per_cap_list[1]
    capacity_per_wind_40_50 = per_cap_list[2]
    
    capacity_per_wind_list = [capacity_per_wind_20_29] * 10 + [capacity_per_wind_30_39] * 10 + [capacity_per_wind_40_50] * 11   
 
    future_c_list = future_inflow_on.copy() * 1000 # convert to kW
    #future_n_list = future_c_list / capacity_per_wind   
    future_n_list = [future_c_list[i] / capacity_per_wind_list[i] for i in range(len(future_c_list))] # number of wind turbines
    future_d_list = get_diameter(np.array(capacity_per_wind_list))  # diameter of wind turbines
    future_h_list = get_height(np.array(capacity_per_wind_list))  # height of wind turbines
    
    oh_dict = load_onshore_dict(path=excel_path)
    print(oh_dict)
    
    c_list, d_list, h_list, nacl_list, tower_list, time_list = load_original_data(path=excel_path)
    hist_mass_by_year = calculate_material_mass_by_year(c_list, d_list, h_list, nacl_list, tower_list, time_list, oh_dict)
    
    time_list = [2020 + i for i in range(len(future_n_list))]
    
    future_mass_by_year = calculate_future_material_mass_onshore_by_year(future_n_list, np.array(capacity_per_wind_list), future_d_list, future_h_list, oh_dict, time_list, tp=tp)
    
    mass_by_year = {}
    for year in np.sort(list(hist_mass_by_year.keys())):
        mass_by_year[year] = hist_mass_by_year[year]
    for year in future_mass_by_year:
        mass_by_year[year] = future_mass_by_year[year]
    
    # save to csv
    df = pd.DataFrame(mass_by_year).T
    # sort using the index
    df = df.sort_index()
    df.to_csv('results/material_onshore_mass_by_year_{}.csv'.format(tp))
    plot_mass_by_year(df, 'save_figs/material_onshore_{}.pdf'.format(tp), w_scale=6)

    return mass_by_year, ratio_on


if __name__ == '__main__':
    capacity_onshore(tp=1)