import argparse
import json
from collections.abc import Iterable
from datetime import datetime
from math import ceil
from pathlib import Path
from typing import Any

import tabulate
from openpyxl.cell.cell import Cell

tabulate.PRESERVE_WHITESPACE = True

type_contents = dict[str, dict[str, int | float]]  # contents, feedstock
type_preforms = dict[str, dict[str, int | float | str]]
type_subitems = dict[str, list[str]]

dumps_dir = Path('json')
calls_dir = Path('calls')
for folder in dumps_dir, calls_dir:
    if not folder.exists():
        folder.mkdir()

now_as_string = lambda: datetime.now().strftime('%Y-%m-%d %H-%M-%S')

holy_tabs = 'contents', 'preforms', 'feedstock', 'subitems'
paths = [dumps_dir / f'{tab}.json' for tab in holy_tabs]


def read_dump(path: Path) -> type_contents | type_preforms | type_subitems:
    """Read dump file and return a deserialized object. """
    with path.open(encoding='windows-1251') as file:
        obj = json.load(file)
    return obj


if all(path.exists() for path in paths):
    contents, preforms, feedstock, subitems = [read_dump(path) for path in paths]


def get_command(argv: list[str]) -> argparse.Namespace:
    """Parse command line arguments. """
    
    parser = argparse.ArgumentParser(
        formatter_class=argparse.RawTextHelpFormatter,
        description='CALCUMET. Sheet metal calculator',
        epilog='Meredelin Evgeny, meredelin@pm.me, 2022')
    
    parser.add_argument(
        'action', choices=['go', 'call', 'update'], 
        help="Action to perform.\n\n"
             "go: \tCalculate number of metal sheets required to fulfill " 
             "given items demand.\n"
             "\tMay be followed by optional GROUPED parameter.\n\n"
             "call: \tSearch for items containing any substring from "
             "tester list and represent\n\titems' parameters.\n\n"
             "update: Instantiate/update json dump and loader's tab with "
             "respective processor func.\n\n")
    
    parser.add_argument(
        'targets', nargs='*', 
        help="0+ list containing tester for CALL or dump(s) for UPDATE.\n")
    
    parser.add_argument(
        '-g', '--grouped', nargs='*', 
        help="0+ list of substrings filtering items that may share a metal"
             " sheet.")
    
    command = parser.parse_args(argv)
    return command.action, command.targets, command.grouped


def sort_nested_dict(d: dict[Any, dict]) -> dict[Any, dict]:
    """Sort nested dictionary. """
    for key, val in d.items():
        d[key] = dict(sorted(val.items()))
    return dict(sorted(d.items()))


def get_cells_values(cells: Iterable[Cell]) -> list[str | int | float]:
    """Get a list of openpyxl cells values. """
    return [cell.value for cell in cells]


def get_pps(a: int | float, b: int | float, 
            sheet: str, feedstock: type_contents) -> int:
    """Calculate the best preforms per sheet number. """
    length, width = feedstock[sheet]['L'], feedstock[sheet]['W']
    return max((length // a) * (width // b), (length // b) * (width // a))


def get_ceil_delta(value: float | int) -> float | int:
    """Get difference between ceiled value and value. """
    return ceil(value) - value


def spread_delta(d: dict[Any, float | int]) -> dict[Any, float | int]:
    """Spread the delta between sum and its, the sum, ceil across 
    summands starting from the one closest to its, the summand, ceil. 
    """
    
    f = lambda item: get_ceil_delta(item[1])
    d = dict(sorted(d.items(), key=f))
    spread = get_ceil_delta(sum(d.values()))
    
    for key, val in d.items():
        delta = get_ceil_delta(val)
        
        # x = (delta, spread)[delta > spread]
        # d[key] += x
        # spread -= x
        # if not spread:
        #     break
        
        if spread < delta:
            d[key] += spread
            break
        d[key] = ceil(val)
        spread -= delta
        if not spread:
            break    
    
    return d


def ceil_dict_values(d: dict[Any, float | int]) -> dict[Any, int]:
    """Ceil dictionary values. """
    return {key: ceil(val) for key, val in d.items()}


transformer = {
    'grouped': spread_delta, 
    'single': ceil_dict_values
}


def has_marker(item: str, tester: Iterable[str]) -> bool:
    """Test if a string contains any word from tester collection. """
    return any(word in (item, item.lower())[word.islower()]
               for word in tester)
    
    
def print_table(table: list[list[str]], headers: list[str]) -> None:
    """Print a table with tabulate module. """
    width = [12] * len(headers)
    print(tabulate.tabulate(table, headers=headers, tablefmt='grid',
                            maxheadercolwidths=width, maxcolwidths=width))
    