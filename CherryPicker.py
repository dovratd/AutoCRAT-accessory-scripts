
""" RNSA cherry picker """

#
# Create AutoCRAT-style 'Rep Summary' and 'RNSA' files
# for a selected subpopulation of cells.
#


import itertools
from pathlib import Path

import pandas as pd
from AutoCRAT_RepTime import export_rep_summary
from AutoCRAT_RNSA import create_rnsa_summary, export_rnsa

# Additional dependencies: openpyxl, xlsxwriter


""" Parameters """


# Folder in which all the relevant files are located.
folder = r''
# Name of the AutoCRAT "Rep Summary" file to be cherry-picked.
# If a RNSA file is located in the same folder with an identical name
# (except with "RNSA" instead of "Rep Summary"), it will also undergo cherry-picking.
rep_summary_filename = ''
# Name of an Excel file with the list of cells to select. This file should be
# formatted with position names in the first column and cell numbers in the second
# (as the "Clustered" file created by the RNSAheatmap script).
selected_cells_filename = ''

# Parameters for RNSA summary chart.
colors = ['red', 'orange', 'lime']
rnsa_x_axis = [-2, 3]
rnsa_y_axis = [0.1, 0.8]


""" Functions """


def read_files(folder, selected_cells_filename, rep_summary_filename, rnsa_filename):
    """
    Read Excel files
    """

    # Read the list of desired cells.
    selected_cells = pd.read_excel(Path(folder, selected_cells_filename),
                                   sheet_name=0,
                                   header=0,
                                   usecols='A:B')
    selected_cells.columns = ['Field', 'Cell']

    # Read the replication summary.
    rep_summary = pd.read_excel(Path(folder, rep_summary_filename),
                                sheet_name='Summary',
                                header=0,
                                index_col=0,
                                keep_default_na=False)
    # Discard unneeded columns.
    rep_summary = rep_summary.loc[:, [c for c in rep_summary.columns if 'Unnamed' not in c]]
    # This is just to put the field name in all the rows of the table for convenience.
    for row_num, row_value in rep_summary['Field'].items():
        if not row_value:
            rep_summary.at[row_num, 'Field'] = rep_summary.at[row_num - 1, 'Field']

    # Read the first 3 sheets of the RNSA file, if available.
    try:
        rnsa = pd.read_excel(Path(folder, rnsa_filename),
                             sheet_name=None,
                             header=[0, 1],
                             index_col=0)
    except FileNotFoundError:
        rnsa = None

    return selected_cells, rep_summary, rnsa


def screened_rep_summary(selected_cells, rep_summary, folder, rep_summary_filename, screened_string):
    """
    Create new replication summary file screened according to selection
    """

    # Create a copy of selected_cells with "n" instead of "Cell_n".
    selected_cells2 = selected_cells.copy()
    selected_cells2['Cell'] = selected_cells2['Cell'].str.split('_', expand=True)[1].astype(int)
    # Create new replication summary, screened by intersection with the selected cells.
    rep_summary_screened = pd.merge(rep_summary, selected_cells2, on=['Field', 'Cell'], how='inner')
    # In Excel, cell numbering is 1-indexed.
    rep_summary_screened.index += 1
    # Remove lines in the replication summary that lack data for deltaT. This can happen
    # due to multiple lines having the same cell number (different cycles).
    rep_summary_screened = rep_summary_screened[rep_summary_screened.iloc[:, 5] != '']

    # Extract channel names, deltaT designations and deltaT ranges
    # from the summary table headers.
    c_names = [c for c in rep_summary.columns if
               c != 'Field' and
               c != 'Cell' and
               '->' not in c and
               c != 'DSB']
    delta_t_names = {c_pair: 'deltaT_' + c_pair[0] + '->' + c_pair[1]
                     for c_pair in itertools.combinations(c_names, 2)}
    delta_t_range = set(tuple(c.split('[')[1].split(']')[0].split(', ')) for c in rep_summary.columns if '[' in c)
    delta_t_range = [int(x) if float(x).is_integer() else float(x) for x in list(delta_t_range)[0]]

    # Export merged replication summary to Excel.
    new_summary_path = Path(folder, rep_summary_filename.replace('.xlsx', screened_string))
    export_rep_summary(rep_summary_screened, new_summary_path, c_names, delta_t_names, delta_t_range)


def screened_rnsa(selected_cells, rnsa, folder, rnsa_filename, screened_string):
    """
    Create new RNSA file screened according to selection
    """

    rnsa_channels = list(rnsa.keys())[:3]

    new_rnsa = {}
    # Create new RNSA tables, screened by intersection with the selected cells.
    for c_name in rnsa_channels:

        new_rnsa[c_name] = pd.DataFrame(
            index=rnsa[c_name].index,
            columns=pd.MultiIndex.from_tuples(selected_cells.itertuples(index=False, name=None))
        )
        new_rnsa[c_name] = rnsa[c_name][rnsa[c_name].columns.intersection(new_rnsa[c_name].columns)]

    # Create new RNSA summary table.
    new_rnsa['Summary'] = create_rnsa_summary(new_rnsa, rnsa_channels)

    # Export new RNSA to Excel.
    new_rnsa_path = Path(folder, rnsa_filename.replace('.xlsx', screened_string))
    colors_dict = {k: v for k, v in zip(rnsa_channels, colors)}
    export_rnsa(new_rnsa, new_rnsa_path, rnsa_channels, colors_dict, rnsa_x_axis, rnsa_y_axis)


""" Main script """


if __name__ == '__main__':

    rep_summary_filename, selected_cells_filename = (
        n + '.xlsx' for n in (rep_summary_filename, selected_cells_filename)
        if '.xlsx' not in n
    )
    rnsa_filename = rep_summary_filename.replace('Rep Summary', 'RNSA')
    screened_string = ' - Screened by ' + selected_cells_filename.split('.')[0] + '.xlsx'

    # Read relevant Excel files.
    selected_cells, rep_summary, rnsa = read_files(
        folder, selected_cells_filename, rep_summary_filename, rnsa_filename
    )

    # Create new replication summary file screened according to selection.
    screened_rep_summary(selected_cells, rep_summary, folder, rep_summary_filename, screened_string)

    if rnsa:
        # Create new replication summary file screened according to selection.
        screened_rnsa(selected_cells, rnsa, folder, rnsa_filename, screened_string)
