
""" RNSA sigmoid fitter """

#
# This script takes a RNSA file and fits a sigmoidal function to a defined
# time window on the averaged replisome-normalized signal.
#


from pathlib import Path

import numpy as np
import pandas as pd
from scipy import optimize
from AutoCRAT_RNSA import export_rnsa_summary

# Additional dependencies: openpyxl, xlsxwriter


""" Parameters """


# Location and filename of the relevant RNSA file.
rnsa_folder = r''
rnsa_filename = ''

# Start and end point for the sigmoidal fit.
fit_window_start = -1
fit_window_end = 2

# Name of channel for which to perform fitting.
channel_of_interest = ''

# Parameters for RNSA summary chart.
colors = ['red', 'orange', 'lime']
rnsa_x_axis = [-2, 3]
rnsa_y_axis = [0.1, 0.8]


""" Functions """


def logistic(x, base, height, steepness, midpoint):
    """
    Definition of the logistic function, the sigmoidal function used to fit the data
    """

    return base + height / (1 + np.exp(-steepness * (x - midpoint)))


def fit_sigmoid(data_to_fit):
    """
    Fit the sigmoidal function to the data using SciPy
    """

    initial_guess = [data_to_fit.min(),
                     data_to_fit.max() - data_to_fit.min(),
                     1,
                     np.median(data_to_fit.index)
                     ]

    popt, pcov = optimize.curve_fit(logistic, data_to_fit.index, data_to_fit, p0=initial_guess, nan_policy='omit')

    return popt


def export_sigmoid_rnsa(rnsa_folder, new_rnsa_filename, rnsa_channels, rnsa_summary, midpoint):
    """
    Create new RNSA file with the fit
    """

    writer = pd.ExcelWriter(Path(rnsa_folder, new_rnsa_filename), engine='xlsxwriter')

    # Export RNSA summary table to Excel and create RNSA chart.
    colors_dict = {k: v for k, v in zip(rnsa_channels, colors)}
    chart = export_rnsa_summary(rnsa_summary, writer, rnsa_channels, colors_dict, rnsa_x_axis, rnsa_y_axis)

    # Draw the fitted sigmoid on the RNSA chart.
    max_row = rnsa_summary.shape[0] + 2
    chart.add_series({
        'categories': ['RNSA_Summary', 3, 0, max_row, 0],
        'values': ['RNSA_Summary', 3, 10, max_row, 10],
        'line': {
            'color': 'black',
            'width': 3,
            'dash_type': 'dash'
        },
        'name': 'Sigmoid',
    })
    # Write the midpoint of the fit.
    writer.sheets['RNSA_Summary'].write(4, 28,
                                        'Midpoint in window ' +
                                        str(fit_window_start) + ' - ' +
                                        str(fit_window_end) + ':')
    writer.sheets['RNSA_Summary'].write(4, 31, midpoint)

    writer.close()


""" Main script """


if __name__ == '__main__':

    if '.xlsx' not in rnsa_filename:
        rnsa_filename = rnsa_filename + '.xlsx'

    # Get the existing RNSA Excel file.
    rnsa_excel = pd.ExcelFile(Path(rnsa_folder, rnsa_filename))
    # Get the channel names from the names of the first 3 sheets of the Excel file.
    rnsa_channels = rnsa_excel.sheet_names[:3]
    # Read the RNSA Summary.
    rnsa_summary = rnsa_excel.parse(sheet_name='RNSA_Summary',
                                    header=[0, 1],
                                    index_col=0)

    # Select the data to fit from the Mean column of the selected channel,
    # bound by the start and end points of the desired time window.
    data_to_fit = rnsa_summary.loc[fit_window_start:fit_window_end, (channel_of_interest, 'Mean')]

    # Fit the sigmoidal function to the data using SciPy.
    fit = fit_sigmoid(data_to_fit)

    # Create a theoretical sigmoidal curve based on the results of the fit.
    # Add it to the RNSA Summary table, so it will be exported to the new file.
    fitted_curve = pd.Series(index=data_to_fit.index)
    fitted_curve[:] = logistic(fitted_curve.index, *fit)
    rnsa_summary.loc[fitted_curve.index, (channel_of_interest, 'Sigmoid')] = fitted_curve

    # Create new RNSA file with the fit.
    new_rnsa_filename = rnsa_filename.replace('.xlsx', ' - Sigmoid.xlsx')
    export_sigmoid_rnsa(rnsa_folder, new_rnsa_filename, rnsa_channels, rnsa_summary, fit[3])
