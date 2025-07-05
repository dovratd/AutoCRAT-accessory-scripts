
""" Create heatmap from RNSA results """

#
# This script takes a RNSA file and displays the results
# of all individual cells as a clustered heatmap.
#


from pathlib import Path

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from scipy import cluster

# Additional dependencies: openpyxl, xlsxwriter


""" Parameters """


# Location and filename of the relevant RNSA file.
rnsa_folder = r''
rnsa_filename = ''

# Desired color of heatmap.
heatmap_color = 'lime'

# Lower and upper bounds of the X axis to display.
# This is the 'normalized time' axis in RNSA charts.
# Note: this will also affect clustering results, since clustering
# is performed only on the desired segment of the data.
rnsa_x_axis = [-2, 3]

# Omit cells with a high proportion of missing data.
# Cells for which too much data is missing within these bounds
# will be omitted from the heatmap. This range must be smaller
# than or equal to the range defined in rnsa_x_axis.
max_nan_range = [0, 3]
# The maximum allowed percentage of missing data within the above bound.
max_nans = 30


""" Functions """


def cluster_df(df):
    """
    Perform clustering analysis on DataFrame using SciPy
    """

    # Turn NaNs (missing data) into zeroes.
    # This is necessary for clustering and also useful for heatmap display.
    # Warning: this assumes missing data means no signal, which is a problematic
    # assumption! Beware when relying on the heatmap as representation of real data!
    clustered_df = df.fillna(0)

    # Cluster using the 'average' method.
    z = cluster.hierarchy.linkage(clustered_df.T, method='average')
    # Flatten dendrogram.
    # Maximizing cluster number (by using 'maxclust' with the number of cells
    # as upper bound) was found to produce the most informative heatmaps.
    clusters = cluster.hierarchy.fcluster(z, df.shape[1], criterion='maxclust')

    # Sort DataFrame by clusters (by adding a 'Cluster' row,
    # sorting by it and deleting it).
    clustered_df.loc['Cluster', :] = clusters
    clustered_df = clustered_df.sort_values(by=['Cluster'], axis=1)
    clustered_df = clustered_df.drop(index='Cluster')

    return clustered_df


def create_heatmap(df):
    """
    Create heatmap using Seaborn
    """

    ax = sns.heatmap(df.T,
                     cbar=False,
                     cmap=sns.dark_palette(heatmap_color, as_cmap=True),
                     xticklabels=round(
                         df.shape[0] /
                         (len(set(np.linspace(rnsa_x_axis[0], rnsa_x_axis[1], 20).astype('int'))) - 1)
                     ),
                     yticklabels=1)
    ax.set_ylabel('')

    # Save heatmap as PNG file.
    ax.tick_params(axis='x', labelsize=5)
    ax.tick_params(axis='y', labelsize=1)
    plt.savefig(
        Path(rnsa_folder, rnsa_filename.replace('.xlsx', ' - Heatmap.png')),
        dpi=1200,
        bbox_inches='tight'
    )

    # Display heatmap.
    ax.tick_params(axis='x', labelsize=15)
    ax.tick_params(axis='y', labelsize=4)
    plt.show()


""" Main script """


if __name__ == '__main__':

    if '.xlsx' not in rnsa_filename:
        rnsa_filename = rnsa_filename + '.xlsx'

    # Read the existing RNSA file.
    # Read only the third sheet, which contains the signal being replisome-normalized.
    rnsa_data = pd.read_excel(Path(rnsa_folder, rnsa_filename),
                              sheet_name=2,
                              header=[0, 1],
                              index_col=0)

    # Lower the resolution of the RNSA data along the normalized time axis
    # by averaging each 10 consecutive time points.
    chunky_rnsa = rnsa_data.groupby(np.round(rnsa_data.index, 2), sort=False).mean()
    # Take only the desired segment of the normalized time axis.
    chunky_rnsa = chunky_rnsa.loc[rnsa_x_axis[0]:rnsa_x_axis[1]]
    # Omit cells with a percentage of nans higher than the defined threshold,
    # within the defined range.
    chunky_rnsa = chunky_rnsa[
        chunky_rnsa
        .loc[:,
             chunky_rnsa.loc[max_nan_range[0]:max_nan_range[1]].isna().sum(axis=0) /
             chunky_rnsa.loc[max_nan_range[0]:max_nan_range[1]].shape[0] <
             max_nans / 100]
        .columns
    ]

    # Perform clustering analysis on RNSA data using SciPy.
    clustered_rnsa = cluster_df(chunky_rnsa)

    # Write clustered heatmap to Excel.
    new_filename = rnsa_filename.replace('.xlsx', ' - Clustered.xlsx')
    writer = pd.ExcelWriter(Path(rnsa_folder, new_filename), engine='xlsxwriter')
    clustered_rnsa.T.to_excel(
        writer,
        sheet_name='Clustered RNSA',
        float_format="%.3f",
        freeze_panes=(1, 2),
        merge_cells=False
    )
    writer.sheets['Clustered RNSA'].autofit()
    writer.close()

    # Create heatmap using Seaborn.
    create_heatmap(clustered_rnsa)
