
""" RNSA Subpopulation selector """

#
# After running AutoCRAT in RNSA mode and screening for strong signals using
# ScreenRNSA, this script takes the positive cells and breaks them up into
# subpopulations using defined intensity thresholds during defined time windows.
#


import itertools
from pathlib import Path

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from AutoCRAT_RepTime import export_rep_summary
from AutoCRAT_RNSA import create_rnsa_summary, export_rnsa

# Additional dependencies: openpyxl, xlsxwriter


""" Parameters """


# Location and filename of the relevant RNSA file.
folder = r''
rnsa_filename = ''

time_windows = ([-0.5, 0], [0.6, 1.4], [2.2, 3])
pop1_thresholds = (0.35, 0.3, 0.4)
pop2_thresholds = (0.35, 0.4, 0.3)

# The maximum allowed percentage of missing data in the examined time windows.
max_nans = 75

# Names of the subpopulations.
pop_names = ['Early', 'Late', 'Other']

# Parameters for RNSA summary chart.
colors = ['red', 'orange', 'lime']
rnsa_x_axis = [-1, 3]
rnsa_y_axis = [0.1, 0.8]


""" Functions """


def read_summary_files(folder, rep_summary_filename, rnsa_filename):
    """
    Read replication summary and RNSA files from Excel
    """

    # Read data from the replication summary file.
    rep_summary = pd.read_excel(Path(folder, rep_summary_filename),
                                sheet_name='Positive',
                                header=0,
                                index_col=0)
    # Put the field name in each row of the table, in case it
    # only appears in the first row of each field.
    for row_num, row_value in rep_summary['Field'].items():
        if pd.isna(row_value):
            rep_summary.at[row_num, 'Field'] = rep_summary.at[row_num - 1, 'Field']
    # Discard unneeded columns.
    rep_summary = rep_summary.loc[:, [c for c in rep_summary.columns if 'Unnamed' not in c]]
    # Extract channel names, deltaT designations and deltaT ranges
    # from the summary table headers.
    c_names = [c for c in rep_summary.columns if
               c != 'Field' and
               c != 'Cell' and
               '->' not in c and
               c != 'DSB' and
               c != 'Average dot intensity']
    delta_t_names = {c_pair: 'deltaT_' + c_pair[0] + '->' + c_pair[1]
                     for c_pair in itertools.combinations(c_names, 2)}

    # Read existing RNSA file from Excel.
    rnsa = pd.read_excel(Path(folder, rnsa_filename),
                         sheet_name=None,
                         header=[0, 1],
                         index_col=0)
    rnsa_channels = tuple(rnsa.keys())[:3]

    return rep_summary, c_names, delta_t_names, rnsa, rnsa_channels


def select_subpops(rnsa):
    """
    Assign each cell to a subpopulation
    """

    cell_lists = {name: [] for name in pop_names}
    for (field, cell) in rnsa['EGFP'].columns:

        cell_data = {}
        for i, window in enumerate(time_windows):

            cell_data[i] = rnsa['EGFP'].loc[:, (field, cell)].loc[window[0]:window[1]]

        nan_threshold = 1 - (max_nans / 100)
        if (cell_data[0].mean() < pop1_thresholds[0] and
                cell_data[0].count() / cell_data[0].size > nan_threshold and
                cell_data[1].mean() > pop1_thresholds[1] and
                cell_data[1].count() / cell_data[1].size > nan_threshold and
                cell_data[2].mean() < pop1_thresholds[2] and
                cell_data[2].count() / cell_data[2].size > nan_threshold):

            cell_lists[pop_names[0]].append((field, cell))

        elif (cell_data[0].mean() < pop2_thresholds[0] and
                cell_data[0].count() / cell_data[0].size > nan_threshold and
                cell_data[1].mean() < pop2_thresholds[1] and
                cell_data[1].count() / cell_data[1].size > nan_threshold and
                cell_data[2].mean() > pop2_thresholds[2] and
                cell_data[2].count() / cell_data[2].size > nan_threshold):

            cell_lists[pop_names[1]].append((field, cell))

        elif (cell_data[0].count() / cell_data[0].size > nan_threshold and
                cell_data[1].count() / cell_data[1].size > nan_threshold and
                cell_data[2].count() / cell_data[2].size > nan_threshold):

            cell_lists[pop_names[2]].append((field, cell))

    cell_lists['All'] = [v for sl in cell_lists.values() for v in sl]

    return cell_lists


def subpop_rep_summary(subpop_cell_lists, old_rep_summary, c_names, delta_t_names):
    """
    Create and export replication summaries for each subpopulation
    """

    subpop_rep_summaries = {}
    for name in subpop_cell_lists.keys():
        # Keep the old replication summary, but only where it
        # intersects with the subpopulation cell list.
        subpop_rep_summaries[name] = old_rep_summary.merge(
            pd.DataFrame(
                [(f, int(c.split('_')[1])) for (f, c) in subpop_cell_lists[name]],
                columns=['Field', 'Cell'])
        )
        subpop_rep_summaries[name].index += 1

        # Export replication summary of each subpopulation to Excel.
        subpop_summary_filename = rep_summary_filename.replace('.xlsx', ' - ' + name + '.xlsx')
        subpop_summary_path = Path(folder, subpop_summary_filename)
        export_rep_summary(subpop_rep_summaries[name], subpop_summary_path, c_names, delta_t_names)


def subpop_rnsa(subpop_cell_lists, old_rnsa, rnsa_channels):
    """
    Create and export RNSA files for each subpopulation
    """

    subpop_rnsas = {name: {} for name in subpop_cell_lists.keys()}
    for name in subpop_rnsas.keys():
        # Keep the old RNSA, but only where it intersects
        # with the subpopulation cell list.
        for c in rnsa_channels:
            subpop_rnsas[name][c] = old_rnsa[c].loc[:, subpop_cell_lists[name]]

        # Create new RNSA summary table.
        subpop_rnsas[name]['Summary'] = create_rnsa_summary(subpop_rnsas[name], rnsa_channels)

        # Export RNSA of each subpopulation to Excel.
        subpop_rnsa_filename = rnsa_filename.replace('.xlsx', ' - ' + name + '.xlsx')
        subpop_rnsa_path = Path(folder, subpop_rnsa_filename)
        colors_dict = {k: v for k, v in zip(rnsa_channels, colors)}
        export_rnsa(subpop_rnsas[name], subpop_rnsa_path, rnsa_channels, colors_dict, rnsa_x_axis, rnsa_y_axis)

    return subpop_rnsas


def create_heatmap(subpop_rnsas, rnsa_channels):
    """
    Create heatmap of all subpopulations using Seaborn and export to Excel
    """

    chunky_rnsas = {}
    for name in pop_names:
        relevant_rnsa = subpop_rnsas[name][rnsa_channels[2]]
        # Lower the resolution of the RNSA data along the normalized time axis
        # by averaging each 10 consecutive time points.
        chunky_rnsas[name] = relevant_rnsa.groupby(np.round(relevant_rnsa.index, 2), sort=False).mean()
        # Take only the desired segment of the normalized time axis.
        chunky_rnsas[name] = chunky_rnsas[name].loc[rnsa_x_axis[0]:rnsa_x_axis[1]]
        # Omit cells with a percentage of nans higher than the defined threshold,
        # within the defined range.
        chunky_rnsas[name] = chunky_rnsas[name][
            chunky_rnsas[name]
            .loc[:, chunky_rnsas[name].isna().sum(axis=0) / chunky_rnsas[name].shape[0] < max_nans / 100]
            .columns
        ]

    chunky_rnsa = pd.concat(chunky_rnsas, axis=1)
    ax = sns.heatmap(chunky_rnsa.T,
                     cbar=False,
                     cmap=sns.dark_palette(colors[2], as_cmap=True),
                     xticklabels=round(
                         chunky_rnsa.shape[0] /
                         (len(set(np.linspace(rnsa_x_axis[0], rnsa_x_axis[1], 20).astype('int'))) - 1)
                     ),
                     yticklabels=1)
    ax.set_ylabel('')

    # Save heatmap as PNG file.
    ax.tick_params(axis='x', labelsize=5)
    ax.tick_params(axis='y', labelsize=1)
    plt.savefig(
        Path(folder, rnsa_filename.replace('.xlsx', ' - Subpopulation Heatmap.png')),
        dpi=1200,
        bbox_inches='tight'
    )

    # Display heatmap.
    ax.tick_params(axis='x', labelsize=15)
    ax.tick_params(axis='y', labelsize=4)
    plt.show()

    # Write subpopulation heatmap to Excel.
    new_filename = rnsa_filename.replace('.xlsx', ' - Subpopulations.xlsx')
    writer = pd.ExcelWriter(Path(folder, new_filename), engine='xlsxwriter')
    chunky_rnsa.T.to_excel(
        writer,
        sheet_name='RNSA by subpopulations',
        float_format="%.3f",
        freeze_panes=(1, 3)
    )
    writer.sheets['RNSA by subpopulations'].autofit()
    writer.close()


""" Main script """


if __name__ == '__main__':

    if '.xlsx' not in rnsa_filename:
        rnsa_filename = rnsa_filename + '.xlsx'
    rep_summary_filename = rnsa_filename.replace('RNSA', 'Rep Summary')

    # Read replication summary and RNSA files from Excel.
    old_rep_summary, c_names, delta_t_names, old_rnsa, rnsa_channels = read_summary_files(
        folder, rep_summary_filename, rnsa_filename
    )

    # Assign each cell to a subpopulation.
    subpop_cell_lists = select_subpops(old_rnsa)

    # Create and export replication summaries for each subpopulation.
    subpop_rep_summary(subpop_cell_lists, old_rep_summary, c_names, delta_t_names)

    # Create and export RNSA files for each subpopulation.
    rnsa_data = subpop_rnsa(subpop_cell_lists, old_rnsa, rnsa_channels)

    # Create heatmap of all subpopulations and export to Excel.
    create_heatmap(rnsa_data, rnsa_channels)
