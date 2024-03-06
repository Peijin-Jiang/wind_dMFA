import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from scipy.stats import weibull_min


def get_capacity_data_from_excel(excel_path):
    df_on = pd.read_excel(excel_path, sheet_name='on_capacity')
    df_off = pd.read_excel(excel_path, sheet_name='off_capacity')

    # extract the future onshore stock capacity data
    on_future_capacity_stock_org = df_on['on_future_capacity_stock'].dropna().to_list()
    off_future_capacity_stock_org = df_off['off_future_capacity'].dropna().to_list()

    on_stock = {
        'SSP5': on_future_capacity_stock_org[8: 15],
        'SSP1': on_future_capacity_stock_org[16: 23],
        'SSP4': on_future_capacity_stock_org[24: 31],
        'SSP3': on_future_capacity_stock_org[32:],
        'SSP2': on_future_capacity_stock_org[0: 7],
    }

    on_future_year = df_on['on_future_year'].dropna().to_list()[0: 7]
    on_future_year = [int(year) for year in on_future_year]
    on_historical_year = df_on['on_historical_year'].dropna().to_list()
    on_historical_year = [int(year) for year in on_historical_year]
    on_historical_capacity_inflow = df_on['on_historical_capacity_inflow'].dropna().values

    # extract the future offshore stock capacity data
    off_stock = {
        'SSP5': off_future_capacity_stock_org[8: 15],
        'SSP1': off_future_capacity_stock_org[16: 23],
        'SSP4': off_future_capacity_stock_org[24: 31],
        'SSP3': off_future_capacity_stock_org[32: 39],
        'SSP2': off_future_capacity_stock_org[0: 7],
    }

    off_future_year = df_off['off_future_year'].dropna().to_list()[0: 7]
    off_future_year = [int(year) for year in off_future_year]

    return on_stock, on_future_year, on_historical_year, on_historical_capacity_inflow, off_stock, off_future_year

def capacity_flow_smooth(tp=0):
    # load the excel file and read the data
    excel_path = "input_data/Wind_data.xls"

    # read the lifetime data
    lifetimes = pd.read_excel(excel_path, sheet_name='tech_dev', header=1)
    lifetimes = list(lifetimes['lifetime_summary'].dropna().values)

    # read the capacity data from the excel
    on_stock, on_future_year, on_historical_year, on_historical_capacity_inflow, off_stock, off_future_year = get_capacity_data_from_excel(excel_path)
    
    # plot and save onshore and offshore data
    fig, axs = plt.subplots(2, 3, figsize=(18, 10), constrained_layout=True)

    # different wind energy demand scenarios
    mode_label = {'SSP2': 'SSP2', 'SSP5': 'SSP5', 'SSP1': 'SSP1', 'SSP4': 'SSP4', 'SSP3': 'SSP3'}
    modes_color = {'SSP2': '#006d2c', 'SSP5': '#fc8d59', 'SSP1': '#fee8c8', 'SSP4': '#fdd49e', 'SSP3': '#fdbb84'}
    modes = ['SSP5', 'SSP3', 'SSP1', 'SSP4', 'SSP2']

    for mode in modes:

        on_future_capacity_stock = on_stock[mode]
        off_future_capacity = off_stock[mode]

        # interpolate future stock for onshore and offshore
        annual_years_on = np.arange(on_future_year[0], on_future_year[-1] + 1)
        annual_years_off = np.arange(off_future_year[0], off_future_year[-1] + 1)
        interpolated_future_stock_on = np.interp(annual_years_on, on_future_year, on_future_capacity_stock)
        interpolated_future_stock_off = np.interp(annual_years_off, off_future_year, off_future_capacity)

        # generate random lifetimes for onshore turbines based on weibull distribution
        def get_weibull_scale_from_mean(mean, shape):

            from scipy.special import gamma

            scale = mean / gamma(1 + 1/shape)
            return scale
        
        # assume the shape parameter is 2
        # read the average lifetime from excel (20, 25 or 30 years)
        shape_parameter_on_historical = 2
        average_on_historical = lifetimes[0] 
        scale_parameter_on_historical = get_weibull_scale_from_mean(average_on_historical, shape_parameter_on_historical)
        mean, var, skew, kurt = weibull_min.stats(shape_parameter_on_historical, scale=scale_parameter_on_historical, moments='mvsk')
        print(mean, var, skew, kurt)

        shape_parameter_on_future = 2
        average_on_future = lifetimes[tp]
        scale_parameter_on_future = get_weibull_scale_from_mean(average_on_future, shape_parameter_on_future)   

        shape_parameter_off = 2
        average_off = lifetimes[tp]
        scale_parameter_off = get_weibull_scale_from_mean(average_off, shape_parameter_off)

        random_lifetimes_on_historical = np.round(np.random.weibull(shape_parameter_on_historical, len(on_historical_year)) * scale_parameter_on_historical)
        random_lifetimes_on_future = np.round(np.random.weibull(shape_parameter_on_future, len(annual_years_on)) * scale_parameter_on_future)
        random_lifetimes_off = np.round(np.random.weibull(shape_parameter_off, len(annual_years_off)) * scale_parameter_off)


        # compute Weibull CDF for a range of ages
        max_age_on_historical = int(np.max(random_lifetimes_on_historical))
        ages_on_historical = np.arange(0, max_age_on_historical + 1)
        weibull_cdf_on_historical = weibull_min.cdf(ages_on_historical, shape_parameter_on_historical, scale=scale_parameter_on_historical)
        max_age_on_future = int(np.max(random_lifetimes_on_future))
        ages_on_future = np.arange(0, max_age_on_future + 1)
        weibull_cdf_on_future = weibull_min.cdf(ages_on_future, shape_parameter_on_future, scale=scale_parameter_on_future)
        max_age_off = int(np.max(random_lifetimes_off))
        ages_off = np.arange(0, max_age_off + 1)
        weibull_cdf_off = weibull_min.cdf(ages_off, shape_parameter_off, scale=scale_parameter_off)

        # cohort for turbines 
        historical_outflow_on = np.zeros(len(on_historical_year))
        historical_stock_on = np.zeros(len(on_historical_year))
        cohorts_on_historical = np.zeros((len(on_historical_year), max_age_on_historical + 1))

        future_inflow_on = np.zeros(len(annual_years_on))
        future_stock_on = np.zeros(len(annual_years_on))
        future_outflow_on = np.zeros(len(annual_years_on))
        total_years = len(on_historical_year) + len(annual_years_on)
        cohorts_on_all_his = np.zeros((total_years, max_age_on_historical + 1))
        cohorts_on_all_fut= np.zeros((total_years, max_age_on_future + 1))

        future_inflow_off = np.zeros(len(annual_years_off))
        future_stock_off = interpolated_future_stock_off  
        future_outflow_off = np.zeros(len(annual_years_off))
        total_years_off = len(annual_years_off)
        cohorts_off = np.zeros((total_years_off, max_age_off + 1))
        
        out_flow_contribution_on = np.zeros([total_years, total_years]) # for exampe, [1, 10] means the contribution of year 1 to year 10
        out_flow_contribution_off = np.zeros([total_years_off, total_years_off]) # for exampe, [1, 10] means the contribution of year 1 to year 10

        all_years_on = np.concatenate((on_historical_year, annual_years_on))
        all_years_off = annual_years_off.copy()
        
        # populate the initial cohort sizes
        for i, inflow in enumerate(on_historical_capacity_inflow):
            cohorts_on_historical[i, 0] = inflow

        # iterate through each year and calculate outflow based on Weibull CDF
        for i, year in enumerate(on_historical_year):
            for j in range(i+1):
                cohort_age = year - on_historical_year[j]
                if 0 <= cohort_age < max_age_on_historical:
                    # calculate the outflow based on the Weibull CDF
                    failure_rate = weibull_cdf_on_historical[cohort_age] - weibull_cdf_on_historical[cohort_age - 1] if cohort_age > 0 else 0
                    outflow = cohorts_on_historical[j, 0] * failure_rate
                    historical_outflow_on[i] += outflow
                    # update the contribution of each year to the outflow
                    out_flow_contribution_on[j, i] = outflow
                    # update the remaining capacity
                    last_year_remaining = cohorts_on_historical[j, cohort_age - 1] if cohort_age > 0 else cohorts_on_historical[j, 0]
                    cohorts_on_historical[j, cohort_age] = last_year_remaining - outflow

                # update the stock for the year
                # cohorts on historical means the remaining inflow of ith year at jth age
                # The diagonal is the stock for each year
                historical_stock_on[i] += np.sum(cohorts_on_historical[j, cohort_age])


        # full the historical data
        # for i in range(len(on_historical_year)):
            # cohorts_on_all_his[i, 0] = on_historical_capacity_inflow[i]
        cohorts_on_all_his[:len(on_historical_year), :] = cohorts_on_historical.copy()

        # iterate for the future years
        for i in range(len(on_historical_year), total_years):
            year = annual_years_on[i - len(on_historical_year)]
            
            # calculate the difference between the interpolated future stock
            future_inflow_on[i - len(on_historical_year)] = (interpolated_future_stock_on[i - len(on_historical_year)] - \
                                    (interpolated_future_stock_on[i - len(on_historical_year) - 1] if i > len(on_historical_year) else historical_stock_on[-1]))

            # calculate the outflow and inflow for the future years
            if i == len(on_historical_year): # calculat the inflow of 2020
                hist_on_fut0_outflow = 0
                for j in range(len(on_historical_year)):
                    cohort_age = int(year - on_historical_year[j])
                    if 0 <= cohort_age < max_age_on_historical:
                        failure_rate = weibull_cdf_on_historical[cohort_age] - weibull_cdf_on_historical[cohort_age-1] if cohort_age > 0 else 0
                        outflow = cohorts_on_all_his[j, 0] * failure_rate
                        hist_on_fut0_outflow += outflow
                        out_flow_contribution_on[j, i] = outflow
                future_inflow_on[i - len(on_historical_year)] = future_inflow_on[i - len(on_historical_year)] + hist_on_fut0_outflow
                future_outflow_on[i - len(on_historical_year)] = hist_on_fut0_outflow
            else:
                prev_on_fut_outflow = 0
                for j in range(len(on_historical_year)): #calculate the outflow from historical inflow
                    cohort_age = int(year - on_historical_year[j])
                    if 0 <= cohort_age < max_age_on_historical:
                        failure_rate = weibull_cdf_on_historical[cohort_age] - weibull_cdf_on_historical[cohort_age-1] if cohort_age > 0 else 0
                        outflow = cohorts_on_all_his[j, 0] * failure_rate
                        prev_on_fut_outflow += outflow
                        out_flow_contribution_on[j, i] = outflow
                for j in range(i - len(on_historical_year)): #calculate the outflow from future inflow
                    cohort_age = int(year - annual_years_on[j])
                    if 0 <= cohort_age < max_age_on_future:
                        failure_rate = weibull_cdf_on_future[cohort_age] - weibull_cdf_on_future[cohort_age-1] if cohort_age > 0 else 0
                        outflow = cohorts_on_all_fut[len(on_historical_year) + j, 0] * failure_rate
                        prev_on_fut_outflow += outflow
                        out_flow_contribution_on[len(on_historical_year) + j, i] = outflow
                # calculate the future inflow
                future_inflow_on[i - len(on_historical_year)] = future_inflow_on[i - len(on_historical_year)] + prev_on_fut_outflow
                # calculate the future outflow
                future_outflow_on[i - len(on_historical_year)] = prev_on_fut_outflow     
            cohorts_on_all_fut[i, 0] = future_inflow_on[i - len(on_historical_year)]
        
        # combine the historical and future data
        full_years_on = np.concatenate((on_historical_year, annual_years_on))
        full_inflow_on = np.concatenate((on_historical_capacity_inflow, future_inflow_on))
        full_stock_on = np.concatenate((historical_stock_on, interpolated_future_stock_on))
        full_outflow_on = np.concatenate((historical_outflow_on, future_outflow_on))

        # first year inflow is equal to stock
        future_inflow_off[0] = future_stock_off[0]

        # populate the initial cohort sizes for future offshore turbines
        cohorts_off[0, 0] = future_inflow_off[0]

        # calculate future outflow for offshore turbines using age-cohort method
        for i, year in enumerate(annual_years_off):

            if i > 0:
                future_inflow_off[i] = future_stock_off[i] - future_stock_off[i-1]
                prev_on_fut_outflow = 0
                for j in range(len(annual_years_off)):
                    cohort_age = int(year - annual_years_off[j])
                    if 0 <= cohort_age < max_age_off:
                        failure_rate = weibull_cdf_off[cohort_age] - weibull_cdf_off[cohort_age-1] if cohort_age > 0 else 0
                        outflow = cohorts_off[j, 0] * failure_rate
                        prev_on_fut_outflow += outflow
                        out_flow_contribution_off[j, i] = outflow
                future_inflow_off[i] = future_inflow_off[i] + prev_on_fut_outflow
                future_outflow_off[i] = prev_on_fut_outflow
            cohorts_off[i, 0] = future_inflow_off[i]
                    
            # update the remaining capacity for future cohorts
            if i < total_years_off - 1:
                cohorts_off[i + 1, 0] = future_inflow_off[i]


        full_years_off = np.concatenate(([0] * len(on_historical_year), annual_years_off))
        full_inflow_off = np.concatenate(([0] * len(on_historical_year), future_inflow_off))
        full_stock_off = np.concatenate(([0] * len(on_historical_year), future_stock_off))
        full_outflow_off = np.concatenate(([0] * len(on_historical_year), future_outflow_off))

        # set the years range from 1993 to 2050 for both onshore and offshore
        years_range = np.arange(1993, 2051)
        years_range_offshore = years_range[years_range >= 2020]  # Offshore starts from 2020

        axs[0, 0].plot(years_range, full_inflow_on, label='Inflow Onshore' + ' ({})'.format(mode_label[mode]),color=modes_color[mode])
        axs[0, 0].set_title('Onshore Inflow')
        axs[0, 0].set_xlabel('Year')
        axs[0, 0].set_ylabel('Capacity')

        axs[0, 1].plot(years_range, full_stock_on, label='Stock Onshore' + ' ({})'.format(mode_label[mode]),color=modes_color[mode])
        axs[0, 1].set_title('Onshore Stock')
        axs[0, 1].set_xlabel('Year')
        axs[0, 1].set_ylabel('Capacity')

        axs[0, 2].plot(years_range, full_outflow_on, label='Outflow Onshore' + ' ({})'.format(mode_label[mode]),color=modes_color[mode])
        axs[0, 2].set_title('Onshore Outflow')
        axs[0, 2].set_xlabel('Year')
        axs[0, 2].set_ylabel('Capacity')

        axs[1, 0].plot(years_range_offshore, full_inflow_off[len(full_inflow_off)-len(years_range_offshore):], label='Inflow Offshore' + ' ({})'.format(mode_label[mode]),color=modes_color[mode])
        axs[1, 0].set_title('Offshore Inflow')
        axs[1, 0].set_xlabel('Year')
        axs[1, 0].set_ylabel('Capacity')

        axs[1, 1].plot(years_range_offshore, full_stock_off[len(full_stock_off)-len(years_range_offshore):], label='Stock Offshore' + ' ({})'.format(mode_label[mode]),color=modes_color[mode])
        axs[1, 1].set_title('Offshore Stock')
        axs[1, 1].set_xlabel('Year')
        axs[1, 1].set_ylabel('Capacity')

        axs[1, 2].plot(years_range_offshore, full_outflow_off[len(full_outflow_off)-len(years_range_offshore):], label='Outflow Offshore' + ' ({})'.format(mode_label[mode]),color=modes_color[mode])
        axs[1, 2].set_title('Offshore Outflow')
        axs[1, 2].set_xlabel('Year')
        axs[1, 2].set_ylabel('Capacity')

        # export the plot data
        onshore_plot_data = {'years': years_range, 'inflow_on': full_inflow_on, 'stock_on': full_stock_on, 'outflow_on': full_outflow_on,}
        offshore_plot_data = {'years': years_range_offshore, 'inflow_off': full_inflow_off[len(full_inflow_off)-len(years_range_offshore):], 'stock_off': full_stock_off[len(full_stock_off)-len(years_range_offshore):], 'outflow_off': full_outflow_off[len(full_outflow_off)-len(years_range_offshore):],}
        # save the plot data
        onshore_plot_data = pd.DataFrame(onshore_plot_data)
        offshore_plot_data = pd.DataFrame(offshore_plot_data)
        os.makedirs('results', exist_ok=True)
        onshore_plot_data.to_csv('results/onshore_plot_data_{}.csv'.format(mode))
        offshore_plot_data.to_csv('results/offshore_plot_data_{}.csv'.format(mode))

        # remove the frames of the legend
        for ax in axs.flat:
            ax.legend(frameon=False)

        # remove the frames of each plot
        for ax in axs.flat:
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
        # the contribution should be divided by the inflow to get the percentage, and then sum
        # expand full_inflow_on [d] to [d, d], each column is the same as full_inflow_on
        full_inflow_on_exp = np.expand_dims(full_inflow_on, axis=1)
        assert full_inflow_on_exp[:, 0].all() == full_inflow_on_exp[:, -1].all()
        
        out_flow_ratio_on = out_flow_contribution_on / (np.expand_dims(full_inflow_on, axis=1) + 1e-100)
        out_flow_ratio_off = out_flow_contribution_off / np.expand_dims(full_inflow_off[len(full_inflow_off)-len(years_range_offshore):], axis=1)
        
        ratio_on = {'ratio': out_flow_ratio_on, 'years': full_years_on}
        ratio_off = {'ratio': out_flow_ratio_off, 'years': full_years_off}
        
        print('Material outflow onshore:', out_flow_ratio_on)
        print('Material outflow offshore:', out_flow_ratio_off)
    
    plt.savefig('save_figs/capacity_flow_{}_all.pdf'.format(tp))
    return future_inflow_on, future_inflow_off, ratio_on, ratio_off

if __name__ == '__main__':
    capacity_flow_smooth()