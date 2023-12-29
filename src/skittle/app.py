import re

import os
import fire
import numpy as np
import pandas as pd
import yaml

PLATE = 'plate'


def load_raw(uri):
    """Load raw spreadsheet values.
    """
    if uri.startswith('drive:'):
        from .drive import Drive
        drive = Drive()
        name = uri[len('drive:'):]
        return drive.get_excel(name, dropna=False, drop_unnamed=False, header=False)
    else:
        raise ValueError(f'URI not supported: {uri}')


def extract_grids(xs):
    """Find possible grids (full or partial plates with title and 
    row/column labels)
    """
    grids = {}
    possible = np.ones(xs.shape, dtype=bool)
    for i, row in enumerate(xs[:-1]):
        for j, _ in enumerate(row):
            if not possible[i, j]:
                continue
            label_is_empty = pd.isnull(xs[i, j])
            cell_below_is_empty = pd.isnull(xs[i+1, j])
            name_is_valid = valid_name(xs[i, j])
            if label_is_empty or not cell_below_is_empty or not name_is_valid:
                continue
            grid = extract_grid(xs, i, j)
            if grid is not None:
                grids[xs[i, j]] = grid
                h, w = grid.shape
                possible[i:i + h + 1, j:j + w] = False
    return grids


def extract_grid(xs, i, j):
    """Attempt to extract grid at row `i` and column `j` in `xs`. If 
    no grid is present, returns None. Otherwise expands grid as much
    as possible.
    """
    grid = xs[i+1:, j:]
    if grid.size == 0:
        return

    height, width = 0, 0
    for label in grid[1:, 0]:
        if not isinstance(label, str):
            break
        if label in 'ABCDEFGHIJKLMNOP':
            height += 1
        else:
            break

    for label in grid[0, 1:]:
        try:
            label = int(label)
            width += 1
        except:
            break

    if height == 0 or width == 0:
        return
    
    return pd.DataFrame(grid[1:height+1, 1:width + 1], 
                 index=pd.Index(grid[1:height+1, 0], name='row'), 
                 columns=pd.Index(grid[0, 1:width+1].astype(int), name='col'))


def pivot_grid(name, df_grid):
    """Convert grid from wide format (rows x columns) into long format. 
    One or more plate labels are extracted from grid name if present.
    """
    name = str(name).strip()
    terms = str(name).split(';')
    df_long = df_grid.stack(dropna=False).rename(terms[0]).reset_index()
    if len(terms) == 2:
        arr = []
        plate_string = terms[1].replace(PLATE, '').strip()
        for plate in plate_string.split(','):
            plate = plate.strip()
            arr += [df_long.assign(**{PLATE: plate})]
        df_long = pd.concat(arr)

    return df_long


def valid_name(x):
    """Validate grid name, allowed formats are "{name}", 
    "{name};plate {i}", "{name};plate {i}, plate {j}"
    """
    x = str(x).strip()
    if not x:
        return False
    if x.count(';') not in (0, 1):
        return False
    return True


@np.vectorize
def is_numeric(x):
    """Numeric type checking for spreadsheet values.
    """
    if pd.isnull(x):
        return False
    try:
        float(str(x))
    except:
        return False
    return True


def extract_maps(xs):
    """Extract named maps from first three columns of `xs`.
    """
    A, B = ~(pd.isnull(xs[:, :2]).T)
    C = is_numeric(xs[:, 2])

    encoded = ''.join([str(x) for x in A + 2*B + 4*C]) + '0'

    maps = {}
    pat = '(076*)(?=0)'
    for match in re.finditer(pat, encoded):
        name = xs[match.start() + 1, 0]
        labels = xs[match.start() + 1:match.end(), 1]
        values = xs[match.start() + 1:match.end(), 2]
        maps[name] = dict(zip(values, labels))
        
    keys = {}
    match = re.match('^(3+)', encoded)
    if match:
        keys = {xs[i, 0]: xs[i, 1] for i in range(match.end())}
    
        
    return keys, maps


def build_table(grids, maps):
    """Should check variables match between grids and maps
    Also check variables are not specified more than once? Default is to keep first
    """
    # convert grids to tables
    longs = [pivot_grid(k, v) for k,v in grids.items()]
    df_all = pd.concat(longs)

    # if no grid names contain plate ID
    if PLATE not in df_all:
        df_all[PLATE] = '1'

    # wells with at least one variable defined
    df_ix = df_all[['row', 'col', PLATE]].drop_duplicates().copy()
    
    # add grid variables one at a time
    index = ['row', 'col', PLATE]
    arr = []
    for x in longs:
        (df_ix
        # add matching variables from this grid
        .merge(x, how='left').dropna()
        # convert to long format (columns are index + ["variable", "value"])
        .melt(index).pipe(arr.append)
        )

    df_all = (pd.concat(arr)
     # return to wide format (one row per well, one column per variable)
     .pivot_table(index=index, columns='variable', values='value', aggfunc='first')
     .reset_index()
     .assign(well=lambda x: x['row'] + x['col'].apply('{:02d}'.format))
     .drop(['row', 'col'], axis=1)
     .set_index(['plate', 'well'])
     .sort_index()
    )
    
    # map indicator variables into labels
    for name in maps:
        df_all[name] = df_all[name].map(maps[name])
    
    # clean up
    df_all.columns.name = None
    df_all = df_all[list(maps)]
    return df_all.dropna(how='all').reset_index()


def load_from_drive_example():
    keys, df_long = parse_skittle_sheet('IS intracellular binders 2022/20220725_tag_test')
    return keys, df_long


def parse_skittle_sheet(uri):
    df_raw = load_raw(uri)
    X = df_raw.values.copy()
    X[X == ''] = None
    grids = extract_grids(X[:, 3:])
    keys, maps = extract_maps(X[:, :3])
    maps = validate_grids_and_maps(grids, maps)
    df_long = build_table(grids, maps)
    return keys, df_long


def validate_grids_and_maps(grids, maps):
    """Check that maps exist for all grids, and return maps
    that are in use.
    """
    maps = maps.copy()
    grid_names = set([x.split(';')[0] for x in grids])
    if x := grid_names - set(maps):
        raise ValueError(f'Grids missing variable definitions: {x}\nVariables are: {maps}')
    if x := set(maps) - grid_names:
        [maps.pop(y) for y in x]
    return maps


def export_sheet(name, output_prefix=''):
    """Export a sheet (local or on google drive).

    :param name: path to sheet, can be local or remote, e.g., 
        "<spreadsheet>/<worksheet>" or "<worksheet>.csv"
    :param output_prefix: prefix for csv and yaml output
    """
    if not os.path.exists(name) and not name.startswith('drive:'):
        name = f'drive:{name}'

    keys, df_long = parse_skittle_sheet(name)
    f = output_prefix + 'wells.csv'
    df_long.to_csv(f, index=None)
    names = ', '.join([x for x in df_long if x not in ('plate', 'well')])

    plate_info = ''
    if 'plate' in df_long:
        plate_names = ', '.join(sorted(set(df_long['plate'])))
        plate_info = f'; plates: {plate_names}'
        
    print(f'Wrote info for {len(df_long)} wells to {f} (variables: {names}{plate_info})')

    f = output_prefix + 'keys.yaml'
    with open(f, 'w') as fh:
        yaml.dump(keys, fh)
    print(f'Wrote info for {len(keys)} keys to {f}')


def list_available_worksheets():
    """Print worksheets visible to the current service account.
    """
    import pygsheets
    from .drive import list_available_sheets
    return list(list_available_sheets())


def main():
    # order is preserved
    commands = [
        'export', 'list'
    ]
    # if the command name is different from the function name
    named = {
        'export': export_sheet,
        'list': list_available_worksheets,
        }

    final = {}
    for k in commands:
        try:
            final[k] = named[k]
        except KeyError:
            final[k] = eval(k)

    try:
        fire.Fire(final)
    except BrokenPipeError:
        pass

if __name__ == '__main__':
    main()