from typing import Any, List, Tuple, Union, Optional, Generator
from openpyxl.worksheet.worksheet import Worksheet

def normalize_range(min_val: int, max_val: Optional[int]) -> Tuple[int, int]:
    """Ensure max is not None and min <= max."""
    if max_val is None:
        max_val = min_val
    return (min_val, max_val) if min_val <= max_val else (max_val, min_val)

def normalize_list(val: Optional[Union[int, List[int]]]) -> List[int]:
    """Ensure the value is returned as a list of integers."""
    if val is None:
        return []

    try:
        iterable = val if isinstance(val, (list, tuple, set, range)) else [val]
        return [int(v) for v in iterable]
    except (ValueError, TypeError):
        raise ValueError(f"Value {val!r} unable to convert to int or list of ints.")
        
def normalize_wrapper(
    min_row: int = 1,
    max_row: Optional[int] = None,
    min_col: int = 1,
    max_col: Optional[int] = None,
    xrows: Optional[Union[int, List[int]]] = None,
    xcols: Optional[Union[int, List[int]]] = None,
    xcoord: Optional[List[Tuple[int, int]]] = None
) -> Tuple[int, int, int, int, List[int], List[int], List[Tuple[int, int]]]:
    norm_min_row, norm_max_row = normalize_range(min_row, max_row)
    norm_min_col, norm_max_col = normalize_range(min_col, max_col)
    norm_xrows = normalize_list(xrows)
    norm_xcols = normalize_list(xcols)
    norm_xcoord = xcoord or []
    
    return (
        norm_min_row, norm_max_row,
        norm_min_col, norm_max_col,
        norm_xrows, norm_xcols, norm_xcoord
    )

def is_excluded(r: int, c: int, xrows: List[int], xcols: List[int], xcoord: List[Tuple[int, int]]) -> bool:
    """Check if the cell is excluded."""
    return r in xrows or c in xcols or (r, c) in xcoord

def byRow(
    ws: Worksheet,
    min_row: int = 1,
    max_row: Optional[int] = None,
    min_col: int = 1,
    max_col: Optional[int] = None,
    xrows: Optional[Union[int, List[int]]] = None,
    xcols: Optional[Union[int, List[int]]] = None,
    xcoord: Optional[List[Tuple[int, int]]] = None,
    fallback: Optional[Any] = None
) -> Generator[List[str], None, None]:
    """Yield rows of values, excluding specified rows, columns, or coordinates."""
    (   min_row, max_row, min_col,
        max_col, xrows, xcols, xcoord
    ) = normalize_wrapper(
        min_row, max_row, min_col,
        max_col, xrows, xcols, xcoord
    )
    
    for row in range(min_row, max_row + 1):
        if row in xrows:
            continue
        row_data = []
        for col in range(min_col, max_col + 1):
            if is_excluded(row, col, xrows, xcols, xcoord):
                continue
            row_data.append(getValue(ws, row, col, fallback))
        yield row_data

def byCol(
    ws: Worksheet,
    min_row: int = 1,
    max_row: Optional[int] = None,
    min_col: int = 1,
    max_col: Optional[int] = None,
    xrows: Optional[Union[int, List[int]]] = None,
    xcols: Optional[Union[int, List[int]]] = None,
    xcoord: Optional[List[Tuple[int, int]]] = None,
    fallback: Optional[Any] = None
) -> Generator[List[str], None, None]:
    """Yield columns of values, excluding specified rows, columns, or coordinates."""
    (   min_row, max_row, min_col,
        max_col, xrows, xcols, xcoord
    ) = normalize_wrapper(
        min_row, max_row, min_col,
        max_col, xrows, xcols, xcoord
    )

    for col in range(min_col, max_col + 1):
        if col in xcols:
            continue
        col_data = []
        for row in range(min_row, max_row + 1):
            if is_excluded(row, col, xrows, xcols, xcoord):
                continue
            col_data.append(getValue(ws, row, col, fallback))
        yield col_data

def getValue(
    ws: Worksheet,
    row: int = 1,
    col: int = 1,
    fallback: Optional[Any] = None
):
    return (
        fallback if (value := ws.cell(row=row, column=col).value) is None
        else (f"{value:,.0f}" if isinstance(value, (int, float)) else value)
    )
    
def getValues(
    ws: Worksheet,
    min_row: int = 1,
    max_row: Optional[int] = None,
    min_col: int = 1,
    max_col: Optional[int] = None,
    xrows: Optional[Union[int, List[int]]] = None,
    xcols: Optional[Union[int, List[int]]] = None,
    xcoord: Optional[List[Tuple[int, int]]] = None,
    fallback: Optional[Any] = None,
    direction: Optional[str] = "row"
) -> List[List[str]]:
    """Return a 2D list of values from the worksheet with exclusions applied."""
    operation = byCol if "col" in str(direction ).lower() else byRow
    return list(operation(ws, min_row, max_row, min_col, max_col, xrows, xcols, xcoord, fallback))



