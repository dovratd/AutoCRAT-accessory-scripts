
""" RNSA plotter """

#
# This script takes a RNSA file and plots the summary using MatPlotLib.
#


from pathlib import Path

import pandas as pd
import matplotlib.pyplot as plt


""" Parameters """


# Location and filename of the relevant RNSA file.
rnsa_folder = r''
rnsa_filename = ''

# Color of each channel (in the order they appear in the RNSA file).
channel_colors = ['red', 'orange', 'lime']
# Relative width of line for each channel.
channel_line_widths = [1.5, 1.5, 3]
# Line transparency for each channel (0-transparent, 1-opaque).
line_transparency = [0.5, 0.5, 1]
# SEM shading transparency for each channel (0-transparent, 1-opaque).
shade_transparency = [0.1, 0.12, 0.25]
# Bottom and top limits on the X axis displayed in the chart.
rnsa_x_axis = [-1, 3]
# Bottom and top limits on the Y axis displayed in the chart.
rnsa_y_axis = [0.2, 0.7]
# Labels to be displayed in the legend. If empty, legend will not be displayed.
legend_labels = ['', '', '']
# Axis names.
rnsa_x_name = 'Replisome-Normalized Time'
rnsa_y_name = 'Normalized Fluorescent Intensity'
# Font for the entire plot.
font = 'Arial'


""" Main script """


if __name__ == '__main__':

    if '.xlsx' not in rnsa_filename:
        rnsa_filename = rnsa_filename + '.xlsx'

    # Read the RNSA Summary from the existing RNSA file.
    rnsa_data = pd.read_excel(Path(rnsa_folder, rnsa_filename),
                              sheet_name='RNSA_Summary',
                              header=[0, 1],
                              index_col=0)

    # Initialize plot.
    plt.rcParams['font.family'] = font
    fig, ax = plt.subplots(figsize=(8, 5))

    # Draw line and SEM (as shading) for each channel.
    for i, c in enumerate(list(dict.fromkeys(rnsa_data.keys().get_level_values(0)))):

        line, = ax.plot(rnsa_data.index,
                        rnsa_data.loc[:, (c, 'Mean')],
                        color=channel_colors[i],
                        linewidth=channel_line_widths[i],
                        alpha=line_transparency[i])

        ax.fill_between(rnsa_data.index,
                        rnsa_data.loc[:, (c, '-SEM')],
                        rnsa_data.loc[:, (c, '+SEM')],
                        color=channel_colors[i],
                        linewidth=0,
                        alpha=shade_transparency[i])

        if legend_labels:
            line.set_label(legend_labels[i])

        # If a sigmoid is present in the RNSA file, draw it too.
        try:
            ax.plot(rnsa_data.index,
                    rnsa_data.loc[:, (c, 'Sigmoid')],
                    color='black',
                    linewidth=channel_line_widths[i] * 0.5,
                    linestyle='--')
        except KeyError:
            pass

    ax.set_xlabel(rnsa_x_name, fontsize='large')
    ax.set_ylabel(rnsa_y_name, fontsize='large')
    ax.set_xlim(rnsa_x_axis[0], rnsa_x_axis[1])
    ax.set_ylim(rnsa_y_axis[0], rnsa_y_axis[1])

    if legend_labels:
        # Place legend in the upper left corner.
        ax.legend(loc=2, edgecolor='black')

    # Save figure.
    plt.savefig(Path(rnsa_folder, rnsa_filename.replace('.xlsx', '.png')),
                dpi=1200,
                format='png',
                bbox_inches='tight')
    #plt.show()
