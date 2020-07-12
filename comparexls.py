"""
Compare two Excel sheets
Inspired by https://pbpython.com/excel-diff-pandas-update.html
For the documentation, download this file and type:
python compare.py --help
"""

import argparse

import pandas as pd
import numpy as np


def report_diff(x):
    """Function to use with groupby.apply to highlight value changes."""
    return x[0] if x[0] == x[1] or pd.isna(x).all() else f'{x[0]} ---> {x[1]}'

def highlight_diff(x):
    """Function to use with style.applymap to highlight value changes."""
    if isinstance(x,str):
        if "--->" in x:
            return 'background-color: orange'
            #return 'background-color: %s' % 'orange'
    return ''

def strip(x):
    """Function to use with applymap to strip whitespaces in a dataframe."""
    return x.strip() if isinstance(x, str) else x


def diff_pd(old_df, new_df, idx_col):
    """Identify differences between two pandas DataFrames using a key column.
    Key column is assumed to have only unique data
    (like a database unique id index)
    Args:
        old_df (pd.DataFrame): first dataframe
        new_df (pd.DataFrame): second dataframe
        idx_col (str|list(str)): column name(s) of the index,
          needs to be present in both DataFrames
    """
    # setting the column name as index for fast operations
    old_df = old_df.set_index(idx_col)
    new_df = new_df.set_index(idx_col)

    # get the added and removed rows
    old_keys = old_df.index
    new_keys = new_df.index
    if isinstance(old_keys, pd.MultiIndex):
        removed_keys = old_keys.difference(new_keys)
        added_keys = new_keys.difference(old_keys)
    else:
        removed_keys = np.setdiff1d(old_keys, new_keys)
        added_keys = np.setdiff1d(new_keys, old_keys)
    out_data = {
        'removed row': old_df.loc[removed_keys],
        'added row': new_df.loc[added_keys]
    }
    # focusing on common data of both dataframes
    common_keys = np.intersect1d(old_keys, new_keys, assume_unique=True)
    common_columns = np.intersect1d(
        old_df.columns, new_df.columns, assume_unique=True
    )
    new_common = new_df.loc[common_keys, common_columns].applymap(strip)
    old_common = old_df.loc[common_keys, common_columns].applymap(strip)
    # get the changed rows keys by dropping identical rows
    # (indexes are ignored, so we'll reset them)
    common_data = pd.concat(
        [old_common.reset_index(), new_common.reset_index()], sort=True
    )
    changed_keys = common_data.drop_duplicates(keep=False)[idx_col]
    if isinstance(changed_keys, pd.Series):
        changed_keys = changed_keys.unique()
    else:
        changed_keys = changed_keys.drop_duplicates().set_index(idx_col).index
    # combining the changed rows via multi level columns
    df_all_changes = pd.concat(
        [old_common.loc[changed_keys], new_common.loc[changed_keys]],
        axis='columns',
        keys=['old', 'new']
    ).swaplevel(axis='columns')
    # using report_diff to merge the changes in a single cell with "-->"
    df_changed = df_all_changes.groupby(level=0, axis=1).apply(
        lambda frame: frame.apply(report_diff, axis=1))

    styleddf = (df_changed.style.applymap(highlight_diff))
    #out_data['changed row'] = df_changed
    out_data['changed row'] = styleddf

    # get the added and removed columns
    old_col = list(old_df.columns)
    new_col = list(new_df.columns)
    removed_col = np.setdiff1d(old_col, new_col)
    added_col = np.setdiff1d(new_col, old_col)

    if(removed_col.size > 0):
        old_diff = old_df.loc[common_keys, removed_col].applymap(strip)
        out_data['removed col'] = old_diff

    if(added_col.size > 0):
        new_diff = new_df.loc[common_keys, added_col].applymap(strip)
        out_data['added col'] = new_diff

    return out_data


def compare_excel(
        path1, sheet1, path2, sheet2, out_path, index_col_name, **kwargs
):
    p1 = path1.split('.')
    p2 = path2.split('.')
    if (p1[1] == 'csv'):
        old_df = pd.read_csv(path1, **kwargs)
    else:
        old_df = pd.read_excel(path1, sheet_name=sheet1, **kwargs)

    if (p2[1] == 'csv'):
        new_df = pd.read_csv(path2, **kwargs)
    else:
        new_df = pd.read_excel(path2, sheet_name=sheet2, **kwargs)

    diff = diff_pd(old_df, new_df, index_col_name)
    with pd.ExcelWriter(out_path) as writer:
        for sname, data in diff.items():
            data.to_excel(writer, engine='openpyxl',sheet_name=sname)
    print(f"Differences saved in {out_path}")

def build_parser():
    cfg = argparse.ArgumentParser(
        description="Compares two excel files and outputs the differences "
                    "in another excel file."
    )
    cfg.add_argument("path1", help="Fist excel file")
    cfg.add_argument("sheet1", help="Name 1st excel sheet to compare.")
    cfg.add_argument("path2", help="Second excel file")
    cfg.add_argument("sheet2", help="Name 2nd excel sheet to compare.")
    cfg.add_argument(
        "key_column",
        help="Name of the column(s) with unique row identifier. It has to be "
             "the actual text of the first row, not the excel notation."
             "Use multiple times to create a composite index.",
        nargs="+",
    )
    cfg.add_argument("-o", "--output-path", default="compared.xlsx",
                     help="Path of the comparison results")
    cfg.add_argument("--skiprows", help='number of rows to skip', type=int,
                     action='append', default=None)
    return cfg


def main():
    cfg = build_parser()
    opt = cfg.parse_args()
    compare_excel(opt.path1, opt.sheet1, opt.path2, opt.sheet2, opt.output_path,
                  opt.key_column, skiprows=opt.skiprows)


if __name__ == '__main__':
    main()