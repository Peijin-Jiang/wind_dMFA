import math
from fig_utils import *
from capacity_flow_all import capacity_flow_smooth
from offshore_material import capacity_offshore

# get environmental impact factor from excel
def get_env_impact(path=''):
    wb = xlrd.open_workbook(path)
    # load existing data
    sheet = wb.sheet_by_name('envir_impact')
    
    env_impact = {}
    
    for i in range(0, 8):
        for j in range(1, 5):
            col_name = sheet.cell_value(0, j)
            if col_name not in env_impact:
                env_impact[col_name] = {}
            row_name = sheet.cell_value(i, 0)
            env_impact[col_name][row_name] = sheet.cell_value(i, j)
            if isinstance(sheet.cell_value(i, j), str) and 'N/A' in sheet.cell_value(i, j):
                env_impact[col_name][row_name] = 0
    return env_impact

# add 10 years impact together
def aggregate_seq(seq, time):
    seq_, time_ = [0, 0, 0], ['2020-2030', '2030-2040', '2040-2050']

    for i in range(len(seq)):
        if time[i] < 2030:
            seq_[0] += seq[i]
        elif time[i] < 2040:
            seq_[1] += seq[i]
        else:
            seq_[2] += seq[i]
    return seq_, time_


en_consume_color = ['#deebf7', '#9ecae1', '#3182bd']
en_save_color = ['#e5f5e0', '#a1d99b', '#31a354']

co2_consume_color = ['#fde0dd', '#fa9fb5', '#c51b8a']
co2_save_color = ['#fff7bc', '#fec44f', '#d95f0e']


if __name__ == '__main__':
    tp=1
    excel_path = "input_data/Wind_data.xls"
    env_impact = get_env_impact(path=excel_path)

    mass_by_year, ratio_off = capacity_offshore(tp=tp)

    time_list = [year for year in mass_by_year]

    # save to csv
    df = pd.DataFrame(mass_by_year).T
    # sort using the index
    df = df.sort_index()

    ratio_off_arr = ratio_off['ratio'] # r[i, j] means the ratio from year i to year j
    out_flow_off_material = np.dot(ratio_off_arr.transpose(), df.values)
    
    for i in range(len(df.columns)):
        mass = df.columns[i]
        print('Material for {} is {}'.format(mass, out_flow_off_material[i]))
    
    path = "input_data/Wind_data.xls"
    table, proc_methods = get_data_from_recy_new(path=path)
    
    results = {}
    for k in table.keys():
        if 'on_shore' in k:
            continue
        
        t_type = k.split('_offshore')[0]
        
        results[t_type] = {}
        
        for i, m in enumerate(df.columns):
            recy = np.asarray(table[k][m]) # [m]
            out_flow_m = out_flow_off_material[:, i] # [n]
            
            out_flow_m = np.dot(out_flow_m.reshape(-1, 1), recy.reshape(1, -1)) # [n, m]
            results[t_type][m] = out_flow_m
            # print outflow m
            print('Outflow for {} is {} on strategy {}'.format(m, out_flow_m, t_type))
    
    # make the figures
    
    en_fig, en_ax = plt.subplots(figsize=(3*FIG_WIDTH, 3*FIG_HEIGHT))
    co2_fig, co2_ax = plt.subplots(figsize=(3*FIG_WIDTH, 3*FIG_HEIGHT))

    strategy_list = results.keys()
    strategy_list = [s for s in strategy_list if 'onshore' not in s]

    for si, sn in enumerate(strategy_list):
        en_consume, en_save = 0, 0
        co2_consume, co2_save = 0, 0

        for i, m in enumerate(mass_by_year[2050]):
            
            mass = np.asarray([mass_by_year[y][m] for y in mass_by_year])
            
            recy_list = results[sn][m][:, 0]

            for k in range(len(recy_list)):
                if recy_list[k] > mass[k]:
                    recy_list[k] = mass[k]
            
            inflow = mass - recy_list
            
            if m == 'Cast Iron':
                m_conv = 'Ferrous metal (steel and iron)'
            elif m == 'Steel':
                m_conv = 'Ferrous metal (steel and iron)'
            elif m == 'Nd' or m == 'Dy':
                m_conv = 'NdFeB (Nd and Dy)'
            elif m == 'Others' or 'Other' in m:
                continue
            else:
                m_conv = m
                
            coef = env_impact['Energy_consumption(MJ kg-1)'][m_conv]
            en_consume += np.array(inflow) * coef
            
            coef = env_impact['Energy_saved (MJ kg-1)'][m_conv]
            en_save += np.array(recy_list) * coef
            
            coef = env_impact['CO2_emission (t)'][m_conv]
            co2_consume += np.array(inflow) * coef
            
            coef = env_impact['CO2_reduction(t)'][m_conv]
            co2_save += np.array(recy_list) * coef
        
        # export env impact to csv
        df = pd.DataFrame({'Energy consumption': en_consume, 'Energy saved': en_save, 'CO2 emission': co2_consume, 'CO2 saved': co2_save}, index=time_list)
        df.to_csv('results/offshore_env_impact_{}.csv'.format(sn))

        en_consume = aggregate_seq(en_consume, time_list)[0]
        en_save = aggregate_seq(en_save, time_list)[0]
        en_net = np.array(en_consume) - np.array(en_save)
        co2_consume = aggregate_seq(co2_consume, time_list)[0]
        co2_save, time_agg = aggregate_seq(co2_save, time_list)
        co2_net = np.array(co2_consume) - np.array(co2_save)
        # plot energy
        
        bar_width = 1 / (len(strategy_list)) * 0.8
        ax = en_ax
        ax.bar(np.arange(len(time_agg)) - si * bar_width, en_consume, color=en_consume_color[si], label='Energy consumption ({})'.format(sn), width=bar_width)
        ax.bar(np.arange(len(time_agg)) - si * bar_width, -np.array(en_save), color=en_save_color[si], label='Energy saved ({})'.format(sn), width=bar_width)
        for j in range(len(en_net)):
            ax.scatter(j - si * bar_width, en_net[j], color='black', marker='*', s=100, zorder=10)
            text = en_net[j]
            ax.text(j - si * bar_width, en_net[j] + 2 + (si + 1) * 1e8/160, f"{text:.2e}", ha='center', va='bottom', fontsize=10)
        ax.set_xticks(np.arange(len(time_agg)))
        ax.set_xticklabels(time_agg)
        ax.set_ylabel('Energy (GJ)')
        ax.legend()

        # plot CO2
        ax = co2_ax
        ax.bar(np.arange(len(time_agg)) - si * bar_width, co2_consume, color=co2_consume_color[si], label='CO2 emission ({})'.format(sn), width=bar_width)
        ax.bar(np.arange(len(time_agg)) - si * bar_width, -np.array(co2_save), color=co2_save_color[si], label='CO2 saved ({})'.format(sn), width=bar_width)
        for j in range(len(co2_net)):
            ax.scatter(j - si * bar_width, co2_net[j], color='black', marker='*', s=100, zorder=10)
            text = co2_net[j]
            # change to scientific notation
            
            ax.text(j - si * bar_width, co2_net[j] + 2 + (si + 1) * 1e7/100, f"{text:.2e}", ha='center', va='bottom', fontsize=10)
        
        ax.set_xticks(np.arange(len(time_agg)))
        ax.set_xticklabels(time_agg)
        ax.set_ylabel('tCO2')
        ax.legend()
        
    # save en_axes to figure
    en_fig.tight_layout()
    en_fig.savefig('save_figs/offshore_energy_{}.pdf'.format(tp))
    
    co2_fig.tight_layout()
    co2_fig.savefig('save_figs/offshore_co2_{}.pdf'.format(tp))
