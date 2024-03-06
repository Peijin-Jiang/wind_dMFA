import os
import math
from tqdm import tqdm
from fig_utils import *
from fig_settings import *
from capacity_flow_all import capacity_flow_smooth
from onshore_material import capacity_onshore


if __name__ == '__main__':
    tp=1
    mass_by_year, ratio_on = capacity_onshore(tp=tp)
    

    df = pd.DataFrame(mass_by_year).T
    # sort using the index
    df = df.sort_index()
    
    ratio_on_arr = ratio_on['ratio'] # r[i, j] means the ratio from year i to year j
    out_flow_on_material = np.dot(ratio_on_arr.transpose(), df.values)
    
    for i in range(len(df.columns)):
        mass = df.columns[i]
        print('Material for {} is {}'.format(mass, out_flow_on_material[i]))
    
    path = "input_data/Wind_data.xls"
    table, proc_methods = get_data_from_recy_new(path=path)
    
    results = {}
    for k in table.keys():
        if 'off_shore' in k:
            continue
        
        t_type = k.split('_onshore')[0]
        
        results[t_type] = {}
        
        for i, m in enumerate(df.columns):
            recy = np.asarray(table[k][m])# [m]
            out_flow_m = out_flow_on_material[:, i] # [n]
            
            out_flow_m = np.dot(out_flow_m.reshape(-1, 1), recy.reshape(1, -1)) # [n, m]
            results[t_type][m] = out_flow_m
            # print outflow m
            print('Outflow for {} is {} on strategy {}'.format(m, out_flow_m, t_type))
    
    df.to_csv('results/material_onshore_mass_by_year_{}.csv'.format(tp))
    
    # three strategies in total, we need to make three figures  EoL_C, EoL_O, EoL100
    for strategy in results:
        
        if 'offshore' in strategy:
            continue
        
        result = results[strategy]
        
        fig, ax = plt.subplots(figsize=(6, 6))
        
        # plot mass by year first
        year = [k for k in mass_by_year]
        materials = [m for m in mass_by_year[year[0]]]
        
        # add every material together
        # mass_by_year_sum = np.zeros(len(year))
        
        # for m in materials:
        #     mass = np.asarray([mass_by_year[y][m] for y in year])
        #     mass_by_year_sum += mass
            
        result_sum = np.zeros_like(result[materials[0]])
        
        for m in materials:
            out_flow_m = result[m]
            result_sum += out_flow_m
        
        #axes[i].bar(year, mass, label=m, color='red')
        
        cum_recy = np.zeros(len(year))
        colors = ['#fbb4ae', '#b3cde3', '#ccebc5', '#decbe4', '#fed9a6']
        for j, p in enumerate(tqdm(proc_methods)):

            #ax.bar(year, result_sum[:, j], bottom=cum_recy, label=p, color=colors[j])

            ax.plot(year, cum_recy + result_sum[:, j], label=p, color=colors[j])

            # fill in color between result_sum[:, j] and cum_recy
            ax.fill_between(year, cum_recy + result_sum[:, j], cum_recy, color=colors[j], edgecolor='none')
            
            cum_recy = result_sum[:, j] + cum_recy
        
        print(result[m][:, 1])

        ax.legend()
        
        ax.spines['right'].set_visible(False)
        ax.spines['top'].set_visible(False)
        
        ax.set_xlabel('Year')
        ax.set_ylabel('Mass [t]')
        
        leg = ax.get_legend()
        leg.get_frame().set_linewidth(0.0)
        
        fig.tight_layout()
            
        plt.savefig('save_figs/onshore_{}_{}.pdf'.format(strategy, tp))

        os.makedirs('results/onshore_EoL', exist_ok=True)
        # export plot data as df
        df = pd.DataFrame(result_sum, columns=proc_methods, index=year)
        df.to_csv('results/onshore_EoL/onshore_{}_sum_{}.csv'.format(strategy, tp))

        for m in result.keys():
            df = pd.DataFrame(result[m], columns=proc_methods, index=year)
            df.to_csv('results/onshore_EoL/onshore_{}_{}_{}.csv'.format(strategy, m, tp))