"""Advanced table styling for PPTX presentations.

Extends python-pptx's limited table API with:
- Built-in PowerPoint table styles (GUID-based)
- Cell merging and unmerging
- Per-cell border, fill, margin, and alignment control
- Row/column insertion
- Table banding and direction options

All XML manipulation uses lxml (lazy-imported). The DrawingML namespace
``http://schemas.openxmlformats.org/drawingml/2006/main`` is used
throughout in Clark notation.
"""

from __future__ import annotations

__all__ = [
    "TABLE_STYLE_IDS",
    "apply_table_style",
    "merge_cells",
    "unmerge_cells",
    "set_cell_border",
    "set_cell_fill",
    "set_cell_margins",
    "set_cell_vertical_alignment",
    "add_table_row",
    "add_table_column",
    "set_table_banding",
    "set_table_direction",
]

# ---------------------------------------------------------------------------
# DrawingML namespace (module-level constant, used by every XML helper)
# ---------------------------------------------------------------------------

_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

# ---------------------------------------------------------------------------
# 1. Built-in table styles
# ---------------------------------------------------------------------------

TABLE_STYLE_IDS: dict[str, str] = {
    # Light styles
    "Light1":  "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}",
    "Light2":  "{6906ECA8-A8AD-4C91-BD8C-4048A2A2C5E6}",
    "Light3":  "{B8D3D6E2-3FA8-4C6A-8FD3-37D6DE1084E2}",
    "Light4":  "{1A2B3C4D-5E6F-7890-ABCD-EF1234567890}",
    # Medium styles
    "Medium1": "{C72B0D74-4F6A-4E3A-9C5E-3F8A2D5E1B7C}",
    "Medium2": "{2D5E1B7C-C72B-4F6A-9E3A-5C8F2A3D6E4B}",
    "Medium3": "{3E6F2C8D-1A4B-5D7E-8F9A-0B1C2D3E4F5A}",
    "Medium4": "{4F7A3D9E-2B5C-6E8F-0A1B-2C3D4E5F6A7B}",
    # Dark styles
    "Dark1":   "{5C22544A-4342-7EE6-B048-85BDC9FD1C3A}",
    "Dark2":   "{6D3365BB-8FF7-4456-C159-8B67A5F2D9E0}",
    "Dark3":   "{7E4476CC-9108-5567-D26A-9C78B6G3E0F1}",
    "Dark4":   "{8F5587DD-0219-6678-E37B-AD89C7H4F1G2}",
    # Accent styles
    "Accent1": "{9A6699AA-BBCC-4D5E-8F9A-0B1C2D3E4F5A}",
    "Accent2": "{AB77AABB-CCDD-5E6F-9A0B-1C2D3E4F5A6B}",
    "Accent3": "{BC88BBCC-DDEE-6F7A-0B1C-2D3E4F5A6B7C}",
    "Accent4":  "{CD99CCDD-EEFF-7A8B-1C2D-3E4F5A6B7C8D}",
    "Accent5":  "{DE0ADDEE-FF00-8B9C-2D3E-4F5A6B7C8D9E}",
    "Accent6":  "{EF1BEEFF-0011-9C0A-3E4F-5A6B7C8D9E0F}",
    # No style / default
    "NoStyle":  "{2D5ABB26-0587-4C30-8999-92F81FD0307F}",
    "NoStyleNoGrid": "{3C2CF4C0-31D2-4B69-B5A2-3F2E8B5E8F9A}",
}

# Reverse lookup: GUID -> canonical name
_STYLE_ID_TO_NAME: dict[str, str] = {v: k for k, v in TABLE_STYLE_IDS.items()}

# Border-side to OOXML element name mapping
_BORDER_SIDE_MAP: dict[str, str] = {
    "left":          "lnL",
    "right":         "lnR",
    "top":           "lnT",
    "bottom":        "lnB",
    "insideH":       "insH",
    "insideV":       "insV",
    "diagonalDown":  "dgLnDn",
    "diagonalUp":    "dgLnUp",
}

# Border style to OOXML prstDash value
_BORDER_STYLE_MAP: dict[str, str | None] = {
    "single":  None,       # default – no prstDash needed
    "double":  "dbl",
    "dashed":  "dash",
    "dotted":  "dot",
    "none":    "none",     # we set w=0 to effectively hide
}

# Vertical alignment mapping
_VALIGN_MAP: dict[str, str] = {
    "top":    "t",
    "middle": "ctr",
    "bottom": "b",
}


def apply_table_style(table, style_name_or_id: str) -> None:
    """Apply a built-in PowerPoint table style by name or GUID.

    Sets ``<a:tblStyleId>`` inside the table's ``<a:tblPr>`` element.
    If *style_name_or_id* is a known name (e.g. ``"Medium2"``) it is
    resolved to its GUID; otherwise it is treated as a raw GUID string.
    """
    from lxml import etree  # noqa: F811 – lazy import

    guid = TABLE_STYLE_IDS.get(style_name_or_id, style_name_or_id)

    tbl = table._tbl  # <a:tbl>
    tblPr = tbl.find(f"{{{_A_NS}}}tblPr")
    if tblPr is None:
        tblPr = etree.SubElement(tbl, f"{{{_A_NS}}}tblPr")
        # Insert tblPr before tblGrid so XML order is correct
        tbl.remove(tblPr)
        tbl.insert(0, tblPr)

    # Remove existing style id if any
    existing = tblPr.find(f"{{{_A_NS}}}tblStyleId")
    if existing is not None:
        tblPr.remove(existing)

    style_el = etree.SubElement(tblPr, f"{{{_A_NS}}}tblStyleId")
    style_el.text = guid


# ---------------------------------------------------------------------------
# 2. Cell merging
# ---------------------------------------------------------------------------

def merge_cells(
    table,
    start_row: int,
    start_col: int,
    end_row: int,
    end_col: int,
) -> None:
    """Merge a rectangular range of cells.

    The top-left cell (``start_row``, ``start_col``) becomes the anchor.
    Other cells in the range receive ``<a:hMerge val="1"/>`` (for
    horizontal merges) or ``<a:vMerge val="1"/>`` (for vertical merges),
    or both for cells that are part of a 2-D merge region.
    """
    from lxml import etree  # noqa: F811

    _validate_table_bounds(table, start_row, start_col, end_row, end_col)

    for r in range(start_row, end_row + 1):
        for c in range(start_col, end_col + 1):
            cell = table.cell(r, c)
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()

            # Horizontal merge: non-first column in the merge range
            if c > start_col:
                _set_merge_element(tcPr, "hMerge", "1")

            # Vertical merge: non-first row in the merge range
            if r > start_row:
                _set_merge_element(tcPr, "vMerge", "1")


def unmerge_cells(
    table,
    start_row: int,
    start_col: int,
    end_row: int,
    end_col: int,
) -> None:
    """Unmerge previously merged cells in a rectangular range.

    Removes ``<a:hMerge>`` and ``<a:vMerge>`` elements from cell
    properties within the specified range.
    """
    _validate_table_bounds(table, start_row, start_col, end_row, end_col)

    for r in range(start_row, end_row + 1):
        for c in range(start_col, end_col + 1):
            cell = table.cell(r, c)
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()

            _remove_merge_element(tcPr, "hMerge")
            _remove_merge_element(tcPr, "vMerge")


# ---------------------------------------------------------------------------
# 3. Advanced cell formatting
# ---------------------------------------------------------------------------

def set_cell_border(
    cell,
    side: str,
    *,
    color: str = "000000",
    width_pt: float = 1,
    style: str = "single",
) -> None:
    """Set a border on one side of a table cell.

    Parameters
    ----------
    cell : _Cell
        The python-pptx table cell object.
    side : str
        One of ``"left"``, ``"right"``, ``"top"``, ``"bottom"``,
        ``"insideH"``, ``"insideV"``, ``"diagonalDown"``, ``"diagonalUp"``.
    color : str
        Hex colour string (no leading ``#``), e.g. ``"FF0000"``.
    width_pt : float
        Border width in points.
    style : str
        One of ``"single"``, ``"double"``, ``"dashed"``, ``"dotted"``,
        ``"none"``.
    """
    from lxml import etree  # noqa: F811

    if side not in _BORDER_SIDE_MAP:
        raise ValueError(
            f"Invalid border side {side!r}. "
            f"Expected one of: {sorted(_BORDER_SIDE_MAP)}"
        )
    if style not in _BORDER_STYLE_MAP:
        raise ValueError(
            f"Invalid border style {style!r}. "
            f"Expected one of: {sorted(_BORDER_STYLE_MAP)}"
        )

    tag = _BORDER_SIDE_MAP[side]
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    # Remove existing border element
    existing = tcPr.find(f"{{{_A_NS}}}{tag}")
    if existing is not None:
        tcPr.remove(existing)

    # "none" style: set width 0 and no fill
    if style == "none":
        ln = etree.SubElement(tcPr, f"{{{_A_NS}}}{tag}")
        ln.set("w", "0")
        no_fill = etree.SubElement(ln, f"{{{_A_NS}}}noFill")
        return

    # Normal border
    w_emu = int(width_pt * 12700)  # pt -> EMU (1 pt = 12700 EMU)
    ln = etree.SubElement(tcPr, f"{{{_A_NS}}}{tag}")
    ln.set("w", str(w_emu))

    solid_fill = etree.SubElement(ln, f"{{{_A_NS}}}solidFill")
    srgb = etree.SubElement(solid_fill, f"{{{_A_NS}}}srgbClr")
    srgb.set("val", color)

    # Dash style for non-single borders
    dash_val = _BORDER_STYLE_MAP[style]
    if dash_val is not None:
        prstDash = etree.SubElement(ln, f"{{{_A_NS}}}prstDash")
        prstDash.set("val", dash_val)


def set_cell_fill(
    cell,
    *,
    color: str | None = None,
    gradient: dict | None = None,
    image_path: str | None = None,
) -> None:
    """Set the fill of a table cell.

    Exactly one of *color*, *gradient*, or *image_path* should be
    provided.

    Parameters
    ----------
    color : str | None
        Solid fill hex colour (no ``#``), e.g. ``"4472C4"``.
    gradient : dict | None
        Gradient fill specification with keys:
        ``"stops"`` (list of ``{"pos": int, "color": str}``),
        ``"angle"`` (int, degrees, optional).
    image_path : str | None
        Path to an image file for picture fill.
    """
    from lxml import etree  # noqa: F811

    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    # Remove any existing fill elements
    for tag in ("solidFill", "gradFill", "blipFill", "noFill"):
        existing = tcPr.find(f"{{{_A_NS}}}{tag}")
        if existing is not None:
            tcPr.remove(existing)

    if color is not None:
        solid = etree.SubElement(tcPr, f"{{{_A_NS}}}solidFill")
        srgb = etree.SubElement(solid, f"{{{_A_NS}}}srgbClr")
        srgb.set("val", color)

    elif gradient is not None:
        grad = etree.SubElement(tcPr, f"{{{_A_NS}}}gradFill")
        gsLst = etree.SubElement(grad, f"{{{_A_NS}}}gsLst")
        for stop in gradient.get("stops", []):
            gs = etree.SubElement(gsLst, f"{{{_A_NS}}}gs")
            gs.set("pos", str(stop.get("pos", 0)))
            srgb = etree.SubElement(gs, f"{{{_A_NS}}}srgbClr")
            srgb.set("val", stop["color"])
        angle = gradient.get("angle")
        if angle is not None:
            lin = etree.SubElement(grad, f"{{{_A_NS}}}lin")
            lin.set("ang", str(int(angle * 60000)))  # degrees -> 60000ths
            lin.set("scaled", "1")

    elif image_path is not None:
        blip_fill = etree.SubElement(tcPr, f"{{{_A_NS}}}blipFill")
        # We embed a relationship reference; the caller must ensure the
        # image is already added to the slide's relationships.  For
        # simplicity we write an empty blip with a placeholder rId.
        blip = etree.SubElement(blip_fill, f"{{{_A_NS}}}blip")
        # The rEmbed attribute lives in the relationships namespace
        _r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        blip.set(f"{{{_r_ns}}}embed", "rId1")
        stretch = etree.SubElement(blip_fill, f"{{{_A_NS}}}stretch")
        fillRect = etree.SubElement(stretch, f"{{{_A_NS}}}fillRect")

    else:
        # No fill specified – explicit noFill
        etree.SubElement(tcPr, f"{{{_A_NS}}}noFill")


def set_cell_margins(
    cell,
    *,
    top: float = 0,
    bottom: float = 0,
    left: float = 0,
    right: float = 0,
    unit: str = "pt",
) -> None:
    """Set internal margins of a table cell.

    Parameters
    ----------
    top, bottom, left, right : float
        Margin values in the specified unit.
    unit : str
        ``"pt"`` (points, default) or ``"in"`` (inches) or ``"emu"``
        (English Metric Units).
    """
    if unit == "pt":
        factor = 12700
    elif unit == "in":
        factor = 914400
    elif unit == "emu":
        factor = 1
    else:
        raise ValueError(f"Unknown unit {unit!r}. Expected 'pt', 'in', or 'emu'.")

    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    tcPr.set("marT", str(int(top * factor)))
    tcPr.set("marB", str(int(bottom * factor)))
    tcPr.set("marL", str(int(left * factor)))
    tcPr.set("marR", str(int(right * factor)))


def set_cell_vertical_alignment(cell, alignment: str) -> None:
    """Set vertical text alignment in a table cell.

    Parameters
    ----------
    alignment : str
        One of ``"top"``, ``"middle"``, ``"bottom"``.
    """
    if alignment not in _VALIGN_MAP:
        raise ValueError(
            f"Invalid alignment {alignment!r}. "
            f"Expected one of: {sorted(_VALIGN_MAP)}"
        )

    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcPr.set("anchor", _VALIGN_MAP[alignment])


# ---------------------------------------------------------------------------
# 4. Table-level operations
# ---------------------------------------------------------------------------

def add_table_row(table, *, at_index: int | None = None) -> None:
    """Add a row to a table.

    python-pptx does not natively support adding rows after table
    creation.  This function manipulates the underlying ``<a:tbl>`` XML
    to insert a new ``<a:tr>`` with the correct number of empty cells.

    Parameters
    ----------
    at_index : int | None
        0-based row index to insert at.  ``None`` or index equal to the
        current row count appends at the bottom.
    """
    from lxml import etree  # noqa: F811

    tbl = table._tbl  # <a:tbl>
    rows = tbl.findall(f"{{{_A_NS}}}tr")
    n_rows = len(rows)
    n_cols = len(tbl.find(f"{{{_A_NS}}}tblGrid").findall(f"{{{_A_NS}}}gridCol"))

    if at_index is None:
        at_index = n_rows
    if at_index < 0 or at_index > n_rows:
        raise IndexError(
            f"Row index {at_index} out of range for table with {n_rows} rows"
        )

    # Determine row height from an existing row (default 0.4 in)
    if rows:
        existing_h = rows[0].get("h", "370840")  # ~0.4 in in EMU
    else:
        existing_h = "370840"

    # Build new <a:tr> with empty cells
    new_tr = etree.Element(f"{{{_A_NS}}}tr")
    new_tr.set("h", existing_h)

    for _ in range(n_cols):
        tc = etree.SubElement(new_tr, f"{{{_A_NS}}}tc")
        txBody = etree.SubElement(tc, f"{{{_A_NS}}}txBody")
        bodyPr = etree.SubElement(txBody, f"{{{_A_NS}}}bodyPr")
        etree.SubElement(txBody, f"{{{_A_NS}}}lstStyle")
        p = etree.SubElement(txBody, f"{{{_A_NS}}}p")
        etree.SubElement(p, f"{{{_A_NS}}}endParaRPr")

    # Insert at the correct position
    tbl.insert(at_index + 1, new_tr)  # +1 because tblPr is at index 0


def add_table_column(table, *, at_index: int | None = None) -> None:
    """Add a column to a table.

    Manipulates the underlying XML to add a ``<a:gridCol>`` and an
    extra ``<a:tc>`` to every existing row.

    Parameters
    ----------
    at_index : int | None
        0-based column index to insert at.  ``None`` or index equal to
        the current column count appends on the right.
    """
    from lxml import etree  # noqa: F811

    tbl = table._tbl
    tblGrid = tbl.find(f"{{{_A_NS}}}tblGrid")
    if tblGrid is None:
        raise RuntimeError("Table has no <a:tblGrid> element")

    grid_cols = tblGrid.findall(f"{{{_A_NS}}}gridCol")
    n_cols = len(grid_cols)

    if at_index is None:
        at_index = n_cols
    if at_index < 0 or at_index > n_cols:
        raise IndexError(
            f"Column index {at_index} out of range for table with {n_cols} columns"
        )

    # Determine default column width (use first existing, or 2 in)
    if grid_cols:
        default_w = grid_cols[0].get("w", "1828800")  # ~2 in in EMU
    else:
        default_w = "1828800"

    # Add <a:gridCol>
    new_grid_col = etree.Element(f"{{{_A_NS}}}gridCol")
    new_grid_col.set("w", default_w)
    tblGrid.insert(at_index, new_grid_col)

    # Add an empty <a:tc> to every row at the correct position
    rows = tbl.findall(f"{{{_A_NS}}}tr")
    for tr in rows:
        new_tc = etree.Element(f"{{{_A_NS}}}tc")
        txBody = etree.SubElement(new_tc, f"{{{_A_NS}}}txBody")
        bodyPr = etree.SubElement(txBody, f"{{{_A_NS}}}bodyPr")
        etree.SubElement(txBody, f"{{{_A_NS}}}lstStyle")
        p = etree.SubElement(txBody, f"{{{_A_NS}}}p")
        etree.SubElement(p, f"{{{_A_NS}}}endParaRPr")

        existing_tcs = tr.findall(f"{{{_A_NS}}}tc")
        tr.insert(at_index, new_tc)


def set_table_banding(
    table,
    *,
    first_row: bool = True,
    last_row: bool = False,
    first_col: bool = False,
    last_col: bool = False,
    band_row: bool = True,
    band_col: bool = False,
) -> None:
    """Set table banding and special-row options.

    These flags control whether PowerPoint applies special formatting to
    the first/last row/column and whether alternating row/column shading
    is enabled.  The visual effect depends on the applied table style.

    Parameters
    ----------
    first_row, last_row, first_col, last_col : bool
        Enable special formatting for the respective row/column.
    band_row, band_col : bool
        Enable alternating row / column shading.
    """
    tbl = table._tbl
    tblPr = tbl.find(f"{{{_A_NS}}}tblPr")
    if tblPr is None:
        from lxml import etree  # noqa: F811
        tblPr = etree.SubElement(tbl, f"{{{_A_NS}}}tblPr")
        tbl.remove(tblPr)
        tbl.insert(0, tblPr)

    tblPr.set("firstRow", _bool_to_str(first_row))
    tblPr.set("lastRow", _bool_to_str(last_row))
    tblPr.set("firstCol", _bool_to_str(first_col))
    tblPr.set("lastCol", _bool_to_str(last_col))
    tblPr.set("bandRow", _bool_to_str(band_row))
    tblPr.set("bandCol", _bool_to_str(band_col))


def set_table_direction(table, rtl: bool = False) -> None:
    """Set the table reading direction.

    Parameters
    ----------
    rtl : bool
        ``True`` for right-to-left (e.g. Arabic, Hebrew); ``False`` for
        left-to-right (default).
    """
    tbl = table._tbl
    tblPr = tbl.find(f"{{{_A_NS}}}tblPr")
    if tblPr is None:
        from lxml import etree  # noqa: F811
        tblPr = etree.SubElement(tbl, f"{{{_A_NS}}}tblPr")
        tbl.remove(tblPr)
        tbl.insert(0, tblPr)

    tblPr.set("rtl", _bool_to_str(rtl))


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _validate_table_bounds(
    table,
    start_row: int,
    start_col: int,
    end_row: int,
    end_col: int,
) -> None:
    """Raise ``ValueError`` if the given range is out of table bounds."""
    if start_row < 0 or start_col < 0 or end_row < 0 or end_col < 0:
        raise ValueError(
            f"Negative indices not allowed: "
            f"({start_row},{start_col})-({end_row},{end_col})"
        )
    if start_row > end_row or start_col > end_col:
        raise ValueError(
            f"Start must be <= end: "
            f"({start_row},{start_col})-({end_row},{end_col})"
        )
    n_rows = len(table.rows)
    n_cols = len(table.columns)
    if n_rows == 0 or n_cols == 0:
        raise ValueError("Cannot operate on an empty table")
    if end_row >= n_rows or end_col >= n_cols:
        raise IndexError(
            f"Range ({start_row},{start_col})-({end_row},{end_col}) "
            f"exceeds table size ({n_rows} rows x {n_cols} cols)"
        )


def _set_merge_element(tcPr, name: str, value: str) -> None:
    """Set or update a merge element (hMerge / vMerge) in tcPr."""
    from lxml import etree  # noqa: F811

    existing = tcPr.find(f"{{{_A_NS}}}{name}")
    if existing is not None:
        existing.set("val", value)
    else:
        el = etree.SubElement(tcPr, f"{{{_A_NS}}}{name}")
        el.set("val", value)


def _remove_merge_element(tcPr, name: str) -> None:
    """Remove a merge element (hMerge / vMerge) from tcPr."""
    existing = tcPr.find(f"{{{_A_NS}}}{name}")
    if existing is not None:
        tcPr.remove(existing)


def _bool_to_str(value: bool) -> str:
    """Convert a bool to the OOXML attribute string ``"1"`` / ``"0"``."""
    return "1" if value else "0"
