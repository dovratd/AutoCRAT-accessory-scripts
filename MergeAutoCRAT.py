
""" Merge results of multiple AutoCRAT runs """

#
# After running AutoCRAT separately on multiple movies, this script merges the
# 'Rep Summary' and 'RNSA' files to generate a summary of all the results.
#


import os
import sys
import argparse
import itertools
from pathlib import Path

import pandas as pd
from AutoCRAT_RepTime import export_rep_summary
from AutoCRAT_RNSA import create_rnsa_summary, export_rnsa

# Additional dependencies: openpyxl, xlsxwriter


def merge_rep_summaries(folders, rep_summary_filenames, merged_folder, merged_filename):
    """
    Read replication summary files, merge and export to Excel
    """

    rep_summaries = []
    for rep_summary_filename in rep_summary_filenames:

        files_to_read = []
        for folder in folders:
            files_to_read.append(
                [os.path.join(folder, i) for i in os.listdir(folder) if rep_summary_filename in i]
            )
        files_to_read = [i for sublist in files_to_read for i in sublist]
        if len(files_to_read) == 0:
            raise ValueError('None of the provided folders contain the file: ', rep_summary_filename)
        elif len(files_to_read) > 1:
            raise ValueError('The provided folders contain more than one file named: ', rep_summary_filename)
        else:
            path_to_read = files_to_read[0]

        # Read data from the replication summary files.
        rep_summaries.append(pd.read_excel(path_to_read,
                                           sheet_name='Summary',
                                           header=0,
                                           index_col=0))
        # Discard unneeded columns.
        rep_summaries[-1] = rep_summaries[-1].loc[:, [c for c in rep_summaries[-1].columns if 'Unnamed' not in c]]

    # Check whether all replication summary tables have identical column headers.
    columns_bool = []
    for rep_summary in rep_summaries[1:]:
        columns_bool.append(all(rep_summary.columns == rep_summaries[0].columns))
    if all(columns_bool):
        # Concatenate all tables into one.
        new_summary = pd.concat(rep_summaries, axis=0, ignore_index=True)
        # In Excel, cell numbering is 1-indexed.
        new_summary.index += 1
    else:
        raise ValueError('The Replication Summary files must have identical column headers!')

    # Extract channel names, deltaT designations and deltaT ranges
    # from the summary table headers.
    c_names = [c for c in new_summary.columns if
               c != 'Field' and
               c != 'Cell' and
               '->' not in c and
               c != 'DSB']
    delta_t_names = {c_pair: 'deltaT_' + c_pair[0] + '->' + c_pair[1]
                     for c_pair in itertools.combinations(c_names, 2)}
    delta_t_range = set(tuple(c.split('[')[1].split(']')[0].split(', ')) for c in new_summary.columns if '[' in c)
    delta_t_range = [int(x) if float(x).is_integer() else float(x) for x in list(delta_t_range)[0]]

    # Export merged replication summary to Excel.
    merged_summary_filename = merged_filename.replace('.xlsx', ' - Rep Summary.xlsx')
    new_summary_path = Path(merged_folder, merged_summary_filename)
    export_rep_summary(new_summary, new_summary_path, c_names, delta_t_names, delta_t_range)


def merge_rnsas(folders, rnsa_filenames, merged_folder, merged_filename, colors, rnsa_x_axis, rnsa_y_axis):
    """
    Read RNSA files, merge and export to Excel
    """

    rnsa_files_to_read = {}
    for rnsa_filename in rnsa_filenames:

        rnsa_files_to_read[rnsa_filename] = []
        for folder in folders:
            rnsa_files_to_read[rnsa_filename].append(
                [os.path.join(folder, i) for i in os.listdir(folder) if rnsa_filename in i]
            )
        rnsa_files_to_read[rnsa_filename] = [i for sublist in rnsa_files_to_read[rnsa_filename] for i in sublist]
        if len(rnsa_files_to_read[rnsa_filename]) > 1:
            raise ValueError('The provided folders contain more than one file named: ', rnsa_filename)
        elif len(rnsa_files_to_read[rnsa_filename]) == 1:
            rnsa_files_to_read[rnsa_filename] = rnsa_files_to_read[rnsa_filename][0]

    num_of_rnsa_files = len([i for i in rnsa_files_to_read.values() if i])
    if num_of_rnsa_files == 0:
        print('No matching RNSA files were found, only Replication Summaries were merged.')
    elif num_of_rnsa_files == 1:
        print('Only one RNSA file was found, so no RNSA merging was performed.')
    elif num_of_rnsa_files > 1:

        rnsas = {}
        rnsa_channels = {}
        for name, path_to_read in rnsa_files_to_read.items():
            if path_to_read:
                # Read data from the RNSA files.
                rnsas[name] = pd.read_excel(path_to_read,
                                            sheet_name=None,
                                            header=[0, 1],
                                            index_col=0)
                # Get channel names from the sheet titles.
                rnsa_channels[name] = tuple(rnsas[name].keys())[:3]

        new_rnsa = {}
        # Merge RNSA tables for each channel. Make sure channel names are identical.
        if len(set(rnsa_channels.values())) == 1:
            rnsa_channels = list(rnsa_channels.values())[0]
            for c_name in rnsa_channels:
                new_rnsa[c_name] = pd.concat([rnsas[filename][c_name] for filename in rnsas],
                                             axis=1, sort=True)

        else:
            raise ValueError('The RNSA files must have identical sheet names!')

        # Create new RNSA summary table.
        new_rnsa['Summary'] = create_rnsa_summary(new_rnsa, rnsa_channels)

        # Export merged RNSA to Excel.
        merged_rnsa_filename = merged_filename.replace('.xlsx', ' - RNSA.xlsx')
        new_rnsa_path = Path(merged_folder, merged_rnsa_filename)
        colors_dict = {k: v for k, v in zip(rnsa_channels, colors)}
        export_rnsa(new_rnsa, new_rnsa_path, rnsa_channels,
                    colors_dict, rnsa_x_axis, rnsa_y_axis)


def main(folders, rep_summary_filenames, merged_folder, merged_filename, colors, rnsa_x_axis, rnsa_y_axis):

    if not folders or not rep_summary_filenames or not merged_filename:
        raise ValueError('Incorrect input! Please double-check arguments.')

    rep_summary_filenames = [filename + '.xlsx' if '.xlsx' not in filename else filename
                             for filename in rep_summary_filenames]
    if '.xlsx' not in merged_filename:
        merged_filename = merged_filename + '.xlsx'

    # Read replication summary files, merge and export to Excel.
    merge_rep_summaries(folders, rep_summary_filenames, merged_folder, merged_filename)

    rnsa_filenames = [filename.replace('Rep Summary', 'RNSA') for filename in rep_summary_filenames]

    # Read RNSA files, merge and export to Excel
    merge_rnsas(folders, rnsa_filenames, merged_folder, merged_filename, colors, rnsa_x_axis, rnsa_y_axis)


if __name__ == '__main__':

    if sys.argv[1:]:
        parser = argparse.ArgumentParser()
        parser.add_argument('-f', '--folders', type=str, nargs='*')
        parser.add_argument('-r', '--rep_summary_filenames', type=str, nargs='*')
        parser.add_argument('-o', '--merged_folder', type=str)
        parser.add_argument('-i', '--merged_filename', type=str)
        parser.add_argument('-c', '--colors', type=str, nargs=3, default=['red', 'orange', 'lime'])
        parser.add_argument('-x', '--rnsa_x_axis', type=float, nargs=2, default=[-2, 3])
        parser.add_argument('-y', '--rnsa_y_axis', type=float, nargs=2, default=[0.1, 0.8])

        args = parser.parse_args()
        main(args.folders, args.rep_summary_filenames, args.merged_folder, args.merged_filename,
             args.colors, args.rnsa_x_axis, args.rnsa_y_axis)

    else:
        # Default arguments with which mergeAutoCRAT will run if it wasn't
        # called from the command line with the arguments provided.
        folders = [
            r'',
            r''
        ]
        rep_summary_filenames = [
            '',
            ''
        ]
        merged_folder = r''
        merged_filename = ''
        colors = ['red', 'orange', 'lime']
        rnsa_x_axis = [-2, 3]
        rnsa_y_axis = [0.1, 0.8]

        main(folders, rep_summary_filenames, merged_folder, merged_filename, colors, rnsa_x_axis, rnsa_y_axis)
