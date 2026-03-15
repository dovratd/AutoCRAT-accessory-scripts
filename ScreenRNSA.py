
""" RNSA cell screener """

#
# After running AutoCRAT in RNSA mode, this script screens the cells which appear
# in the RNSA analysis using a rolling window minimum intensity approach, to
# distinguish cells that actually have a sufficiently bright dot.
#


import os
from pathlib import Path

import pandas as pd
from AutoCRAT_RNSA import create_rnsa_summary, export_rnsa

# Additional dependencies: openpyxl, xlsxwriter


""" Parameters """


# Location and filename of the RNSA file to be screened.
# There must be an AutoCRAT "Rep Summary" file in the same folder,
# and with an identical name (except with "Rep Summary" instead of "RNSA").
rnsa_folder = r''
rnsa_filename = ''

# Folders in which all the relevant AutoCRAT "Results" files are located.
folders = [
    r''
    ]

# Name of channel in which intensity should be examined.
channel_of_interest = ''

# Length of time windows (in timepoints) during which the intensity will be averaged.
window_length = 12
# Minimum intensity (average during time window) which qualifies a cell for selection.
min_intensity = 65

# Parameters for RNSA summary chart.
colors = ['red', 'orange', 'lime']
rnsa_x_axis = [-2, 3]
rnsa_y_axis = [0.1, 0.8]


""" Functions """


def read_summary_files(folder, rep_summary_filename, rnsa_filename):
    """
    Read replication summary and RNSA files from Excel
    """

    # Read data from the replication summary file.
    rep_summary = pd.read_excel(Path(folder, rep_summary_filename),
                                sheet_name='Summary',
                                header=0,
                                index_col=0,
                                keep_default_na=False)
    # This is just to put the field name in all the rows of the table for convenience.
    for row_num, row_value in rep_summary['Field'].items():
        if not row_value:
            rep_summary.at[row_num, 'Field'] = rep_summary.at[row_num - 1, 'Field']

    # Read existing RNSA file from Excel.
    rnsa = pd.read_excel(Path(folder, rnsa_filename),
                         sheet_name=None,
                         header=[0, 1],
                         index_col=0)
    rnsa_channels = tuple(rnsa.keys())[:3]

    # Get the names of the relevant fields/positions included in the RNSA analysis.
    fields = rnsa[rnsa_channels[2]].columns.get_level_values(0).unique().tolist()
    # For each field, get the cells from that field that were included in the RNSA analysis.
    rnsa_cells = {field: [c[1] for c in rnsa[rnsa_channels[2]].columns.tolist() if c[0] == field]
                  for field in fields}

    return rep_summary, rnsa, rnsa_channels, fields, rnsa_cells


def read_results_files(folders, fields, rnsa_cells):
    """
    Read data from the results files
    """

    # Search the folders for AutoCRAT "Results" files that have names which match
    # the names found in the "Field" column of the replication summary file.
    files_to_read = {}
    for field in fields:

        files_to_read[field] = []
        for folder in folders:
            files_to_read[field].append(
                [os.path.join(folder, i) for i in os.listdir(folder) if str(field + ' - Results') in i]
            )
        files_to_read[field] = [i for sublist in files_to_read[field] for i in sublist]

        if len(files_to_read[field]) == 0:
            raise ValueError('The relevant folders do not contain an AutoCRAT Results file for field: ',
                             field)
        elif len(files_to_read[field]) > 1:
            raise ValueError('The relevant folders contain more than one AutoCRAT Results file for field: ',
                             field)
        else:
            files_to_read[field] = files_to_read[field][0]

    data_from_excel = {}
    for field, cells in rnsa_cells.items():
        # Remove duplicate cell names derived from cells with more than
        # one cell cycle (such as 'Cell_1_0').
        cells = list(set([c.split('_')[0] + '_' + c.split('_')[1] for c in cells]))
        # Read data from the relevant results files (only the relevant sheets).
        data_from_excel[field] = pd.read_excel(files_to_read[field],
                                               sheet_name=cells,
                                               header=[0, 1],
                                               index_col=0)

    return data_from_excel


def check_intensity(intensity_timeseries):
    """
    Calculate average intensity over rolling windows
    """

    # Create rolling windows of specified length.
    data_windows = list(intensity_timeseries.rolling(window_length))[(window_length - 1):]
    # Discard time windows in which more than 30% of the timepoints
    # are NaNs, since these are not informative.
    data_windows = [dw for dw in data_windows if sum(dw.isna().values) / len(dw) <= 0.3]
    # If there are relatively few informative time windows, a definitive
    # determination cannot be made, so the cell will not be taken.
    if len(data_windows) / (intensity_timeseries.shape[0] - window_length + 1) < 0.25:
        return False
    else:
        # Check which time windows have an average intensity above the defined threshold.
        windows_above_threshold = [dw for dw in data_windows if dw.mean().values > min_intensity]
        # If at least one time window exists during which the average intensity
        # is above the threshold, this cell will be selected.
        # Return the average intensity of all the relevant time windows.
        if len(windows_above_threshold) > 0:
            return sum([dw.mean().values[0] for dw in windows_above_threshold]) / len(windows_above_threshold)
        else:
            return False


def create_summary_tables(summary_data, relevant_cells, positive_cells):
    """
    Create new summary tables
    """

    # Create two new summary tables, separating the check_intensity positive and negative cells.
    new_tables = {k: pd.DataFrame(columns=summary_data.columns[:5]) for k in [0, 1]}
    for row in summary_data.index:

        current_field = summary_data.loc[row, 'Field']
        current_cell = 'Cell_' + str(summary_data.loc[row, 'Cell'])
        # The new summary tables contain cells from the old summary table that were included
        # in the RNSA analysis. Only rows that have a value at the sixth column (deltaT in range)
        # are selected, in case a single cell has more than one row.
        if current_field in relevant_cells.keys():
            if current_cell in relevant_cells[current_field] and summary_data.loc[row].iloc[5]:

                # Cells are added either to the positive or negative summary table, depending
                # on their value in "positive_cells".
                new_tables[positive_cells[current_field][current_cell]] = pd.concat([
                    new_tables[positive_cells[current_field][current_cell]],
                    pd.DataFrame(summary_data.loc[row][:5]).T
                ], axis=0, ignore_index=True)

    for table in new_tables.values():
        table.index += 1

    return new_tables


def create_summary_file(writer, value_string, summary_table, cell_format, float_format):
    """
    Write new summary table to Excel
    """

    summary_table.to_excel(
        writer,
        sheet_name=value_string,
        float_format="%.3f",
        freeze_panes=(1, 0)
    )

    for field_num, field in enumerate(summary_table['Field'].drop_duplicates()):

        start_merge_range = list(summary_table['Field'].drop_duplicates().index)[field_num]
        try:
            end_merge_range = list(summary_table['Field'].drop_duplicates().index)[field_num + 1] - 1
        except IndexError:
            end_merge_range = summary_table.shape[0]
        writer.sheets[value_string].merge_range(start_merge_range, 1,
                                                end_merge_range, 1,
                                                field,
                                                cell_format)

    writer.sheets[value_string].write(2, summary_table.shape[1] + 2, 'Num. of ' + value_string.lower() + ' cells:')
    writer.sheets[value_string].write(2, summary_table.shape[1] + 6, summary_table.shape[0])
    writer.sheets[value_string].write(3, summary_table.shape[1] + 2, 'Median replication time:')
    if summary_table.shape[0]:
        writer.sheets[value_string].write(3,
                                          summary_table.shape[1] + 6,
                                          summary_table.iloc[:, 4].median(),
                                          float_format)


def screen_old_rnsa(old_rnsa, positive_cells, relevant_cells,
                    rnsa_channels, rnsa_folder, rnsa_filename, screened_string):
    """
    Create new RNSA tables by screening the old RNSA for positive cells and export to Excel
    """

    new_rnsa = {}
    # Create new table for each of the first 3 sheets in the RNSA file.
    # They will contain the same data as the old RNSA sheets, but only the columns
    # corresponding to the cells that came back positive from check_intensity.
    for c_name in rnsa_channels:

        mask = (positive_cells[f][c.split('_')[0] + '_' + c.split('_')[1]]
                if c in relevant_cells[f] else False
                for f, c in old_rnsa[c_name].columns)
        new_rnsa[c_name] = old_rnsa[c_name].loc[:, mask]

    # Create new RNSA summary table.
    new_rnsa['Summary'] = create_rnsa_summary(new_rnsa, rnsa_channels)

    # Export new RNSA to Excel.
    new_rnsa_path = Path(rnsa_folder, rnsa_filename.replace('.xlsx', screened_string))
    colors_dict = {k: v for k, v in zip(rnsa_channels, colors)}
    export_rnsa(new_rnsa, new_rnsa_path, rnsa_channels, colors_dict, rnsa_x_axis, rnsa_y_axis)


""" Main script """


if __name__ == '__main__':

    if '.xlsx' not in rnsa_filename:
        rnsa_filename = rnsa_filename + '.xlsx'
    rep_summary_filename = rnsa_filename.replace('RNSA', 'Rep Summary')

    # Read replication summary and RNSA files from Excel. Create lists of relevant
    # fields and cell names included in the RNSA analysis.
    old_rep_summary, old_rnsa, rnsa_channels, fields, rnsa_cells = read_summary_files(
        rnsa_folder, rep_summary_filename, rnsa_filename
    )

    # Read relevant data from the results files.
    results_data = read_results_files(folders, fields, rnsa_cells)

    # Analyze the intensity timeseries of the 3rd channel using rolling windows of
    # averaged intensity, to check which cells should be included in the new RNSA.
    positive_cells = {}
    average_intensity = {}
    for field, field_data in results_data.items():

        positive_cells[field] = {}
        average_intensity[field] = {}
        for cell_name, cell_data in field_data.items():

            relevant_track_name = [t for t in cell_data.columns.get_level_values(0) if channel_of_interest in t]
            if len(set(relevant_track_name)) == 1 and cell_data.shape[0] >= window_length:
                # Get the intensity data for the relevant track.
                intensity_series = cell_data.loc[:, (relevant_track_name, 'Intensity')]
                # Run the rolling window analysis.
                # Returns average intensity during all the time windows of length
                # 'window_length' that have an average intensity higher than the
                # threshold level, 'min_intensity'.
                # If no windows are above the threshold, returns False.
                above_threshold = check_intensity(intensity_series)
                if above_threshold:
                    positive_cells[field][cell_name] = True
                    average_intensity[field][cell_name] = above_threshold
                else:
                    positive_cells[field][cell_name] = False

            else:
                rnsa_cells[field].remove(cell_name)

    # Create new summary tables.
    new_summary_tables = create_summary_tables(old_rep_summary, rnsa_cells, positive_cells)

    # Add a column to the Positive cell summary containing the average intensity value.
    new_summary_tables[1]['Average dot intensity'] = pd.Series(dtype='float')
    for field_name, field_data in average_intensity.items():
        for cell_name, avg_int in field_data.items():
            new_summary_tables[1].loc[
                (new_summary_tables[1].loc[:, 'Field'] == field_name) &
                (new_summary_tables[1].loc[:, 'Cell'] == int(cell_name.split('_')[1])),
                'Average dot intensity'] = average_intensity[field_name][cell_name]

    # Create new replication summary Excel file.
    screened_string = ' - Screened (Window ' + str(window_length) + ', Int ' + str(min_intensity) + ').xlsx'
    new_summary_filename = rep_summary_filename.replace('.xlsx', screened_string)
    writer = pd.ExcelWriter(Path(rnsa_folder, new_summary_filename), engine='xlsxwriter')

    cell_format = writer.book.add_format({'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
    float_format = writer.book.add_format({'num_format': '0.00'})

    # Write the two new summary tables to Excel.
    for k, table in new_summary_tables.items():

        create_summary_file(writer, ['Negative', 'Positive'][k], table, cell_format, float_format)

    writer.close()

    # Create new RNSA tables by screening the old RNSA for positive cells and export to Excel.
    screen_old_rnsa(old_rnsa, positive_cells, rnsa_cells, rnsa_channels, rnsa_folder, rnsa_filename, screened_string)

