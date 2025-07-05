
""" Screen replication summary and RNSA by absolute time of replication """

#
# After running AutoCRAT, this script screens the cells which appear in the
# replication summary and RNSA files by timepoint of array replication, to
# screen cells that replicate early or late in the movie.
#


import itertools
from pathlib import Path

import pandas as pd
from AutoCRAT_RepTime import export_rep_summary
from AutoCRAT_RNSA import create_rnsa_summary, export_rnsa

# Additional dependencies: openpyxl, xlsxwriter


""" Parameters """


# Location and filename of the "Rep Summary" file to be screened.
# If there is an AutoCRAT "RNSA" file in the same folder and with an identical name
# (except with "RNSA" instead of "Rep Summary"), it will also undergo screening.
folder = r''
rep_summary_filename = ''

# Threshold timepoint beyond which cells will be screened.
rep_time_threshold = 100
# Select which cells to *keep*.
# If 'Under', cells in which all arrays are replicated before the above threshold
# will be kept, cells in which at least one array is replicated at or after the
# threshold will be discarded.
# If 'Over', cells in which all arrays are replicated after the threshold will be
# kept.
under_over = 'under'

# Parameters for RNSA summary chart.
colors = ['red', 'orange', 'lime']
rnsa_x_axis = [-2, 3]
rnsa_y_axis = [0.1, 0.8]


""" Functions """


def read_files(folder, rep_summary_filename, rnsa_filename):
    """
    Read replication summary file and RNSA file from Excel
    """

    # Read data from the replication summary file.
    rep_summary = pd.read_excel(Path(folder, rep_summary_filename),
                                sheet_name='Summary',
                                header=0,
                                index_col=0)
    # This is just to put the field name in all the rows of the table for convenience.
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
               c != 'DSB']
    delta_t_names = {c_pair: 'deltaT_' + c_pair[0] + '->' + c_pair[1]
                     for c_pair in itertools.combinations(c_names, 2)}
    delta_t_range = set(tuple(c.split('[')[1].split(']')[0].split(', ')) for c in rep_summary.columns if '[' in c)
    delta_t_range = [int(x) if float(x).is_integer() else float(x) for x in list(delta_t_range)[0]]

    try:
        # Read data from the RNSA file.
        rnsa = pd.read_excel(Path(folder, rnsa_filename),
                             sheet_name=None,
                             header=[0, 1],
                             index_col=0)

    except FileNotFoundError:
        rnsa = None

    return rep_summary, c_names, delta_t_names, delta_t_range, rnsa


def screen_summary(old_rep_summary, c_names):
    """
    Screen replication summary for cells that are under/over the desired threshold
    """

    new_rep_summary = pd.DataFrame(columns=old_rep_summary.columns)
    remaining_cells = []
    removed_cells = []
    # If any array replication time is under/over the threshold in a given cell,
    # the cell will not be included in the new summary, and will be added to a
    # list of removed cells.
    for row_num in old_rep_summary.index:

        if under_over.casefold() == 'Under'.casefold():
            if all(old_rep_summary.loc[row_num, c_name] < rep_time_threshold or
                   pd.isna(old_rep_summary.loc[row_num, c_name])
                   for c_name in c_names):
                new_rep_summary.loc[row_num] = old_rep_summary.loc[row_num]

                # If the cell is within the screening threshold and has a deltaT
                # within range, keep it in a list for later reference.
                if not old_rep_summary.loc[row_num].isna().iloc[5]:
                    remaining_cells.append((old_rep_summary.loc[row_num, 'Field'],
                                            old_rep_summary.loc[row_num, 'Cell']))

            else:
                removed_cells.append((old_rep_summary.loc[row_num, 'Field'],
                                      old_rep_summary.loc[row_num, 'Cell']))

        elif under_over.casefold() == 'Over'.casefold():
            if all(old_rep_summary.loc[row_num, c_name] > rep_time_threshold or
                   pd.isna(old_rep_summary.loc[row_num, c_name])
                   for c_name in c_names):
                new_rep_summary.loc[row_num] = old_rep_summary.loc[row_num]

                # If the cell is within the screening threshold and has a deltaT
                # within range, keep it in a list for later reference.
                if not old_rep_summary.loc[row_num].isna().iloc[5]:
                    remaining_cells.append((old_rep_summary.loc[row_num, 'Field'],
                                            old_rep_summary.loc[row_num, 'Cell']))

            else:
                removed_cells.append((old_rep_summary.loc[row_num, 'Field'],
                                      old_rep_summary.loc[row_num, 'Cell']))

        else:
            raise ValueError('The under_over parameter must be \'Under\' or \'Over\'!')

    # Cells can only be present in both the remaining and removed cell lists if they
    # have midpoints for more than one cell cycle, some of which are within the
    # screening threshold and some beyond it. In this case, the cell number should
    # not be included in the removal list, so it's RNSA data will be retained.
    removed_cells = [c for c in removed_cells if c not in remaining_cells]

    # Re-index to remove indexing gaps.
    new_rep_summary = new_rep_summary.reset_index(drop=True)
    new_rep_summary.index += 1

    return new_rep_summary, removed_cells


def screen_rnsa(old_rnsa, rnsa_channels, removed_cells, rnsa_filename, screened_string):
    """
    Screen RNSA for cells that are under/over the desired threshold and export to Excel
    """

    new_rnsa = {}
    # Create new table for each of the first 3 sheets in the RNSA file.
    # They will contain the same data as the old RNSA sheets, but only the
    # columns corresponding to the screened cells.
    for c_name in rnsa_channels:

        mask = (
            False
            if (f, int(c.split('_')[1])) in removed_cells
            else True
            for f, c in old_rnsa[c_name].columns
        )
        new_rnsa[c_name] = old_rnsa[c_name].loc[:, mask]

    # Create new RNSA summary table.
    new_rnsa['Summary'] = create_rnsa_summary(new_rnsa, rnsa_channels)

    # Export new RNSA to Excel.
    new_rnsa_path = Path(folder, rnsa_filename.replace('.xlsx', screened_string))
    colors_dict = {k: v for k, v in zip(rnsa_channels, colors)}
    export_rnsa(new_rnsa, new_rnsa_path, rnsa_channels, colors_dict, rnsa_x_axis, rnsa_y_axis)


""" Main script """


if __name__ == '__main__':

    if '.xlsx' not in rep_summary_filename:
        rep_summary_filename = rep_summary_filename + '.xlsx'
    rnsa_filename = rep_summary_filename.replace('Rep Summary', 'RNSA')

    # Read replication summary file from Excel.
    old_rep_summary, c_names, delta_t_names, delta_t_range, old_rnsa = read_files(
        folder, rep_summary_filename, rnsa_filename
    )

    # Screen replication summary for cells that are under/over the desired threshold.
    new_rep_summary, removed_cells = screen_summary(old_rep_summary, c_names)

    # Export merged replication summary to Excel.
    screened_string = ' - Screened (Rep time ' + under_over.lower() + ' ' + str(rep_time_threshold) + ').xlsx'
    new_summary_filename = rep_summary_filename.replace('.xlsx', screened_string)
    new_summary_path = Path(folder, new_summary_filename)
    export_rep_summary(new_rep_summary, new_summary_path, c_names, delta_t_names, delta_t_range)

    if old_rnsa:
        # Screen RNSA for cells that are under/over the desired threshold and export to Excel.
        rnsa_channels = tuple(old_rnsa.keys())[:3]
        screen_rnsa(old_rnsa, rnsa_channels, removed_cells, rnsa_filename, screened_string)
