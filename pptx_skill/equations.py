"""Office Math Markup Language (OMML) equation support for PowerPoint.

Creates ``<m:oMath>`` and ``<m:oMathPara>`` elements that render as native
equations in PowerPoint, Word, and other Office applications.

Supports building equations from a Python API (``add_fraction``,
``add_superscript``, etc.) or converting from LaTeX notation.

OOXML reference: ECMA-376 Part 4, §22.1 (Office Math Markup Language).
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any

__all__ = [
    "add_equation",
    "add_display_equation",
    "add_fraction",
    "add_superscript",
    "add_subscript",
    "add_sub_superscript",
    "add_radical",
    "add_nary",
    "add_delimiter",
    "add_matrix",
    "add_accent",
    "add_bar",
    "add_function",
    "add_group_character",
    "add_limit_lower",
    "add_limit_upper",
    "add_equation_box",
    "add_text_run",
    "latex_to_omml",
    "list_equations",
    "remove_equation",
    "OmmlBuilder",
]

# ---------------------------------------------------------------------------
# Namespace
# ---------------------------------------------------------------------------

_NS_M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
_NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
_NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


# ---------------------------------------------------------------------------
# OMML Builder
# ---------------------------------------------------------------------------

class OmmlBuilder:
    """Fluent builder for OMML equation XML trees.

    Usage::

        eq = OmmlBuilder()
        eq.add_fraction("a", "b")  # a/b
        eq.add_superscript("x", "2")  # x²
        xml = eq.to_xml()  # lxml Element
    """

    def __init__(self):
        from lxml import etree
        self._root = etree.Element(f"{{{_NS_M}}}oMath")

    def add_text_run(self, text: str) -> OmmlBuilder:
        """Add a math text run ``<m:r><m:t>text</m:t></m:r>``."""
        from lxml import etree
        r = etree.SubElement(self._root, f"{{{_NS_M}}}r")
        t = etree.SubElement(r, f"{{{_NS_M}}}t")
        t.text = text
        return self

    def add_fraction(self, numerator: str, denominator: str,
                     *, bar_type: str = "bar") -> OmmlBuilder:
        """Add a fraction ``<m:f>``.

        Parameters
        ----------
        numerator, denominator : str
            Text for numerator and denominator.
        bar_type : str
            "bar" (horizontal bar), "noBar" (no bar), "skw" (skewed).
        """
        from lxml import etree
        f = etree.SubElement(self._root, f"{{{_NS_M}}}f")
        fPr = etree.SubElement(f, f"{{{_NS_M}}}fPr")
        ctrlPr = etree.SubElement(fPr, f"{{{_NS_M}}}ctrlPr")
        type_elem = etree.SubElement(fPr, f"{{{_NS_M}}}type")
        type_elem.set(f"{{{_NS_M}}}val", bar_type)

        num = etree.SubElement(f, f"{{{_NS_M}}}num")
        r = etree.SubElement(num, f"{{{_NS_M}}}r")
        t = etree.SubElement(r, f"{{{_NS_M}}}t")
        t.text = numerator

        den = etree.SubElement(f, f"{{{_NS_M}}}den")
        r = etree.SubElement(den, f"{{{_NS_M}}}r")
        t = etree.SubElement(r, f"{{{_NS_M}}}t")
        t.text = denominator
        return self

    def add_superscript(self, base: str, sup: str) -> OmmlBuilder:
        """Add a superscript ``<m:sSup>``."""
        from lxml import etree
        sSup = etree.SubElement(self._root, f"{{{_NS_M}}}sSup")
        sSupPr = etree.SubElement(sSup, f"{{{_NS_M}}}sSupPr")
        etree.SubElement(sSupPr, f"{{{_NS_M}}}ctrlPr")

        e = etree.SubElement(sSup, f"{{{_NS_M}}}e")
        r = etree.SubElement(e, f"{{{_NS_M}}}r")
        t = etree.SubElement(r, f"{{{_NS_M}}}t")
        t.text = base

        sup = etree.SubElement(sSup, f"{{{_NS_M}}}sup")
        r = etree.SubElement(sup, f"{{{_NS_M}}}r")
        t = etree.SubElement(r, f"{{{_NS_M}}}t")
        t.text = sup
        return self

    def add_subscript(self, base: str, sub: str) -> OmmlBuilder:
        """Add a subscript ``<m:sSub>``."""
        from lxml import etree
        sSub = etree.SubElement(self._root, f"{{{_NS_M}}}sSub")
        sSubPr = etree.SubElement(sSub, f"{{{_NS_M}}}sSubPr")
        etree.SubElement(sSubPr, f"{{{_NS_M}}}ctrlPr")

        e = etree.SubElement(sSub, f"{{{_NS_M}}}e")
        r = etree.SubElement(e, f"{{{_NS_M}}}r")
        t = etree.SubElement(r, f"{{{_NS_M}}}t")
        t.text = base

        sub_elem = etree.SubElement(sSub, f"{{{_NS_M}}}sub")
        r = etree.SubElement(sub_elem, f"{{{_NS_M}}}r")
        t = etree.SubElement(r, f"{{{_NS_M}}}t")
        t.text = sub
        return self

    def add_sub_superscript(self, base: str, sub: str, sup: str) -> OmmlBuilder:
        """Add a sub-superscript ``<m:sSubSup>``."""
        from lxml import etree
        sSubSup = etree.SubElement(self._root, f"{{{_NS_M}}}sSubSup")
        sSubSupPr = etree.SubElement(sSubSup, f"{{{_NS_M}}}sSubSupPr")
        etree.SubElement(sSubSupPr, f"{{{_NS_M}}}ctrlPr")

        e = etree.SubElement(sSubSup, f"{{{_NS_M}}}e")
        r = etree.SubElement(e, f"{{{_NS_M}}}r")
        t = etree.SubElement(r, f"{{{_NS_M}}}t")
        t.text = base

        sub_elem = etree.SubElement(sSubSup, f"{{{_NS_M}}}sub")
        r = etree.SubElement(sub_elem, f"{{{_NS_M}}}r")
        t = etree.SubElement(r, f"{{{_NS_M}}}t")
        t.text = sub

        sup_elem = etree.SubElement(sSubSup, f"{{{_NS_M}}}sup")
        r = etree.SubElement(sup_elem, f"{{{_NS_M}}}r")
        t = etree.SubElement(r, f"{{{_NS_M}}}t")
        t.text = sup
        return self

    def add_radical(self, base: str, degree: str = "",
                    *, hide_degree: bool = False) -> OmmlBuilder:
        """Add a radical (nth root) ``<m:rad>``.

        Parameters
        ----------
        base : str
            Content under the radical sign.
        degree : str
            Degree (empty for square root).
        hide_degree : bool
            If True, hide the degree even if provided.
        """
        from lxml import etree
        rad = etree.SubElement(self._root, f"{{{_NS_M}}}rad")
        radPr = etree.SubElement(rad, f"{{{_NS_M}}}radPr")
        if hide_degree or not degree:
            degHide = etree.SubElement(radPr, f"{{{_NS_M}}}degHide")
            degHide.set(f"{{{_NS_M}}}val", "1")
        etree.SubElement(radPr, f"{{{_NS_M}}}ctrlPr")

        deg = etree.SubElement(rad, f"{{{_NS_M}}}deg")
        if degree:
            r = etree.SubElement(deg, f"{{{_NS_M}}}r")
            t = etree.SubElement(r, f"{{{_NS_M}}}t")
            t.text = degree

        e = etree.SubElement(rad, f"{{{_NS_M}}}e")
        r = etree.SubElement(e, f"{{{_NS_M}}}r")
        t = etree.SubElement(r, f"{{{_NS_M}}}t")
        t.text = base
        return self

    def add_nary(self, base: str, lower: str = "", upper: str = "",
                 *, char: str = "∑") -> OmmlBuilder:
        """Add an n-ary operator (sum, integral, product) ``<m:nary>``.

        Parameters
        ----------
        char : str
            Operator character: ∑ (sum), ∫ (integral), ∏ (product),
            ∐ (coproduct), ∪ (union), ∩ (intersection), ⋃ (big union),
            ⋂ (big intersection), ⊕ (big oplus), etc.
        """
        from lxml import etree
        nary = etree.SubElement(self._root, f"{{{_NS_M}}}nary")
        naryPr = etree.SubElement(nary, f"{{{_NS_M}}}naryPr")
        chr_elem = etree.SubElement(naryPr, f"{{{_NS_M}}}chr")
        chr_elem.set(f"{{{_NS_M}}}val", char)
        etree.SubElement(naryPr, f"{{{_NS_M}}}ctrlPr")

        sub = etree.SubElement(nary, f"{{{_NS_M}}}sub")
        if lower:
            r = etree.SubElement(sub, f"{{{_NS_M}}}r")
            t = etree.SubElement(r, f"{{{_NS_M}}}t")
            t.text = lower

        sup = etree.SubElement(nary, f"{{{_NS_M}}}sup")
        if upper:
            r = etree.SubElement(sup, f"{{{_NS_M}}}r")
            t = etree.SubElement(r, f"{{{_NS_M}}}t")
            t.text = upper

        e = etree.SubElement(nary, f"{{{_NS_M}}}e")
        r = etree.SubElement(e, f"{{{_NS_M}}}r")
        t = etree.SubElement(r, f"{{{_NS_M}}}t")
        t.text = base
        return self

    def add_delimiter(self, *elements: str,
                      begin_char: str = "(", end_char: str = ")",
                      grow: bool = True) -> OmmlBuilder:
        """Add a delimiter (brackets) ``<m:d>``.

        Parameters
        ----------
        *elements : str
            Content elements inside the delimiter.
        begin_char, end_char : str
            Opening and closing characters.
        grow : bool
            Whether delimiters grow to match content height.
        """
        from lxml import etree
        d = etree.SubElement(self._root, f"{{{_NS_M}}}d")
        dPr = etree.SubElement(d, f"{{{_NS_M}}}dPr")
        begChr = etree.SubElement(dPr, f"{{{_NS_M}}}begChr")
        begChr.set(f"{{{_NS_M}}}val", begin_char)
        endChr = etree.SubElement(dPr, f"{{{_NS_M}}}endChr")
        endChr.set(f"{{{_NS_M}}}val", end_char)
        if grow:
            grow_elem = etree.SubElement(dPr, f"{{{_NS_M}}}grow")
            grow_elem.set(f"{{{_NS_M}}}val", "1")
        etree.SubElement(dPr, f"{{{_NS_M}}}ctrlPr")

        for elem_text in elements:
            e = etree.SubElement(d, f"{{{_NS_M}}}e")
            r = etree.SubElement(e, f"{{{_NS_M}}}r")
            t = etree.SubElement(r, f"{{{_NS_M}}}t")
            t.text = elem_text
        return self

    def add_matrix(self, rows: list[list[str]]) -> OmmlBuilder:
        """Add a matrix ``<m:m>``.

        Parameters
        ----------
        rows : list of list of str
            Matrix content as a 2D array of strings.
        """
        from lxml import etree
        m = etree.SubElement(self._root, f"{{{_NS_M}}}m")
        mPr = etree.SubElement(m, f"{{{_NS_M}}}mPr")
        etree.SubElement(mPr, f"{{{_NS_M}}}ctrlPr")

        for row_data in rows:
            mr = etree.SubElement(m, f"{{{_NS_M}}}mr")
            for cell_text in row_data:
                e = etree.SubElement(mr, f"{{{_NS_M}}}e")
                r = etree.SubElement(e, f"{{{_NS_M}}}r")
                t = etree.SubElement(r, f"{{{_NS_M}}}t")
                t.text = cell_text
        return self

    def add_accent(self, base: str, *, char: str = "̂") -> OmmlBuilder:
        """Add an accent (hat, tilde, etc.) ``<m:acc>``.

        Parameters
        ----------
        char : str
            Accent character: ̂ (hat), ̃ (tilde), ̅ (overbar), ̇ (dot),
            ̈ (diaeresis), ̌ (check), ⃗ (arrow), etc.
        """
        from lxml import etree
        acc = etree.SubElement(self._root, f"{{{_NS_M}}}acc")
        accPr = etree.SubElement(acc, f"{{{_NS_M}}}accPr")
        chr_elem = etree.SubElement(accPr, f"{{{_NS_M}}}chr")
        chr_elem.set(f"{{{_NS_M}}}val", char)
        etree.SubElement(accPr, f"{{{_NS_M}}}ctrlPr")

        e = etree.SubElement(acc, f"{{{_NS_M}}}e")
        r = etree.SubElement(e, f"{{{_NS_M}}}r")
        t = etree.SubElement(r, f"{{{_NS_M}}}t")
        t.text = base
        return self

    def add_bar(self, base: str, *, position: str = "top") -> OmmlBuilder:
        """Add an overbar or underbar ``<m:bar>``.

        Parameters
        ----------
        position : str
            "top" for overbar, "bottom" for underbar.
        """
        from lxml import etree
        bar = etree.SubElement(self._root, f"{{{_NS_M}}}bar")
        barPr = etree.SubElement(bar, f"{{{_NS_M}}}barPr")
        pos = etree.SubElement(barPr, f"{{{_NS_M}}}pos")
        pos.set(f"{{{_NS_M}}}val", position)
        etree.SubElement(barPr, f"{{{_NS_M}}}ctrlPr")

        e = etree.SubElement(bar, f"{{{_NS_M}}}e")
        r = etree.SubElement(e, f"{{{_NS_M}}}r")
        t = etree.SubElement(r, f"{{{_NS_M}}}t")
        t.text = base
        return self

    def add_function(self, name: str, argument: str) -> OmmlBuilder:
        """Add a function (e.g., sin, cos, log) ``<m:func>``."""
        from lxml import etree
        func = etree.SubElement(self._root, f"{{{_NS_M}}}func")
        funcPr = etree.SubElement(func, f"{{{_NS_M}}}funcPr")
        etree.SubElement(funcPr, f"{{{_NS_M}}}ctrlPr")

        fName = etree.SubElement(func, f"{{{_NS_M}}}fName")
        r = etree.SubElement(fName, f"{{{_NS_M}}}r")
        t = etree.SubElement(r, f"{{{_NS_M}}}t")
        t.text = name

        e = etree.SubElement(func, f"{{{_NS_M}}}e")
        r = etree.SubElement(e, f"{{{_NS_M}}}r")
        t = etree.SubElement(r, f"{{{_NS_M}}}t")
        t.text = argument
        return self

    def add_group_character(self, base: str, *, char: str = "⏞",
                            position: str = "top") -> OmmlBuilder:
        """Add a group character (brace over/under) ``<m:groupChr>``.

        Parameters
        ----------
        char : str
            Group character: ⏞ (top brace), ⏟ (bottom brace),
            ⏜ (top parenthesis), ⏝ (bottom parenthesis).
        position : str
            "top" or "bottom".
        """
        from lxml import etree
        groupChr = etree.SubElement(self._root, f"{{{_NS_M}}}groupChr")
        groupChrPr = etree.SubElement(groupChr, f"{{{_NS_M}}}groupChrPr")
        chr_elem = etree.SubElement(groupChrPr, f"{{{_NS_M}}}chr")
        chr_elem.set(f"{{{_NS_M}}}val", char)
        pos = etree.SubElement(groupChrPr, f"{{{_NS_M}}}pos")
        pos.set(f"{{{_NS_M}}}val", position)
        etree.SubElement(groupChrPr, f"{{{_NS_M}}}ctrlPr")

        e = etree.SubElement(groupChr, f"{{{_NS_M}}}e")
        r = etree.SubElement(e, f"{{{_NS_M}}}r")
        t = etree.SubElement(r, f"{{{_NS_M}}}t")
        t.text = base
        return self

    def add_limit_lower(self, base: str, limit: str) -> OmmlBuilder:
        """Add a lower limit ``<m:limLow>``."""
        from lxml import etree
        limLow = etree.SubElement(self._root, f"{{{_NS_M}}}limLow")
        limLowPr = etree.SubElement(limLow, f"{{{_NS_M}}}limLowPr")
        etree.SubElement(limLowPr, f"{{{_NS_M}}}ctrlPr")

        e = etree.SubElement(limLow, f"{{{_NS_M}}}e")
        r = etree.SubElement(e, f"{{{_NS_M}}}r")
        t = etree.SubElement(r, f"{{{_NS_M}}}t")
        t.text = base

        lim = etree.SubElement(limLow, f"{{{_NS_M}}}lim")
        r = etree.SubElement(lim, f"{{{_NS_M}}}r")
        t = etree.SubElement(r, f"{{{_NS_M}}}t")
        t.text = limit
        return self

    def add_limit_upper(self, base: str, limit: str) -> OmmlBuilder:
        """Add an upper limit ``<m:limUpp>``."""
        from lxml import etree
        limUpp = etree.SubElement(self._root, f"{{{_NS_M}}}limUpp")
        limUppPr = etree.SubElement(limUpp, f"{{{_NS_M}}}limUppPr")
        etree.SubElement(limUppPr, f"{{{_NS_M}}}ctrlPr")

        e = etree.SubElement(limUpp, f"{{{_NS_M}}}e")
        r = etree.SubElement(e, f"{{{_NS_M}}}r")
        t = etree.SubElement(r, f"{{{_NS_M}}}t")
        t.text = base

        lim = etree.SubElement(limUpp, f"{{{_NS_M}}}lim")
        r = etree.SubElement(lim, f"{{{_NS_M}}}r")
        t = etree.SubElement(r, f"{{{_NS_M}}}t")
        t.text = limit
        return self

    def add_equation_box(self) -> OmmlBuilder:
        """Add a border box around the equation ``<m:borderBox>``."""
        from lxml import etree
        borderBox = etree.SubElement(self._root, f"{{{_NS_M}}}borderBox")
        borderBoxPr = etree.SubElement(borderBox, f"{{{_NS_M}}}borderBoxPr")
        etree.SubElement(borderBoxPr, f"{{{_NS_M}}}ctrlPr")
        e = etree.SubElement(borderBox, f"{{{_NS_M}}}e")
        # The child builder will add to this <m:e>
        return self

    def to_xml(self) -> Any:
        """Return the ``<m:oMath>`` element."""
        return self._root

    def to_display_xml(self, *, justification: str = "center") -> Any:
        """Return a ``<m:oMathPara>`` element (display math)."""
        from lxml import etree
        oMathPara = etree.Element(f"{{{_NS_M}}}oMathPara")
        oMathParaPr = etree.SubElement(oMathPara, f"{{{_NS_M}}}oMathParaPr")
        jc = etree.SubElement(oMathParaPr, f"{{{_NS_M}}}jc")
        jc.set(f"{{{_NS_M}}}val", justification)
        oMathPara.append(self._root)
        return oMathPara


# ---------------------------------------------------------------------------
# Convenience functions that use OmmlBuilder
# ---------------------------------------------------------------------------

def add_fraction(numerator: str, denominator: str) -> Any:
    """Build a fraction OMML element."""
    return OmmlBuilder().add_fraction(numerator, denominator).to_xml()


def add_superscript(base: str, sup: str) -> Any:
    """Build a superscript OMML element."""
    return OmmlBuilder().add_superscript(base, sup).to_xml()


def add_subscript(base: str, sub: str) -> Any:
    """Build a subscript OMML element."""
    return OmmlBuilder().add_subscript(base, sub).to_xml()


def add_sub_superscript(base: str, sub: str, sup: str) -> Any:
    """Build a sub-superscript OMML element."""
    return OmmlBuilder().add_sub_superscript(base, sub, sup).to_xml()


def add_radical(base: str, degree: str = "", **kwargs) -> Any:
    """Build a radical (nth root) OMML element."""
    return OmmlBuilder().add_radical(base, degree, **kwargs).to_xml()


def add_nary(base: str, lower: str = "", upper: str = "", **kwargs) -> Any:
    """Build an n-ary operator OMML element."""
    return OmmlBuilder().add_nary(base, lower, upper, **kwargs).to_xml()


def add_delimiter(*elements: str, **kwargs) -> Any:
    """Build a delimiter (brackets) OMML element."""
    return OmmlBuilder().add_delimiter(*elements, **kwargs).to_xml()


def add_matrix(rows: list[list[str]]) -> Any:
    """Build a matrix OMML element."""
    return OmmlBuilder().add_matrix(rows).to_xml()


def add_accent(base: str, **kwargs) -> Any:
    """Build an accent OMML element."""
    return OmmlBuilder().add_accent(base, **kwargs).to_xml()


def add_bar(base: str, **kwargs) -> Any:
    """Build an overbar/underbar OMML element."""
    return OmmlBuilder().add_bar(base, **kwargs).to_xml()


def add_function(name: str, argument: str) -> Any:
    """Build a function OMML element."""
    return OmmlBuilder().add_function(name, argument).to_xml()


def add_group_character(base: str, **kwargs) -> Any:
    """Build a group character OMML element."""
    return OmmlBuilder().add_group_character(base, **kwargs).to_xml()


def add_limit_lower(base: str, limit: str) -> Any:
    """Build a lower limit OMML element."""
    return OmmlBuilder().add_limit_lower(base, limit).to_xml()


def add_limit_upper(base: str, limit: str) -> Any:
    """Build an upper limit OMML element."""
    return OmmlBuilder().add_limit_upper(base, limit).to_xml()


def add_equation_box() -> Any:
    """Build a border box OMML element."""
    return OmmlBuilder().add_equation_box().to_xml()


def add_text_run(text: str) -> Any:
    """Build a math text run OMML element."""
    return OmmlBuilder().add_text_run(text).to_xml()


# ---------------------------------------------------------------------------
# Insert equation into a slide
# ---------------------------------------------------------------------------

def add_equation(prs_or_path, slide_index: int, *,
                 omml_element=None,
                 left: float = 1.0, top: float = 2.0,
                 width: float = 8.0, height: float = 1.0,
                 name: str | None = None) -> str:
    """Add an inline equation to a text box on a slide.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    omml_element : lxml Element
        An ``<m:oMath>`` element (from OmmlBuilder or convenience functions).
    left, top, width, height : float
        Position and size in inches.
    name : str, optional
        Shape name.

    Returns
    -------
    str
        The shape name.
    """
    from lxml import etree

    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    if slide_index < 1 or slide_index > len(prs.slides):
        raise IndexError(f"slide_index {slide_index} out of range (1..{len(prs.slides)})")
    try:
        slide = prs.slides[slide_index - 1]
        _name = name or f"Equation {len(slide.shapes)}"

        # Create a text box
        txBox = slide.shapes.add_textbox(left, top, width, height)
        txBox.name = _name

        # Clear default paragraph
        tf = txBox.text_frame
        tf.clear()

        # Insert OMML into the paragraph
        p_elem = tf.paragraphs[0]._p

        if omml_element is not None:
            # Register math namespace if not present
            nsmap = p_elem.nsmap if hasattr(p_elem, 'nsmap') else {}
            if _NS_M not in nsmap.values():
                p_elem.append(omml_element)
            else:
                p_elem.append(omml_element)

        return _name
    finally:
        _save_prs(prs, path)


def add_display_equation(prs_or_path, slide_index: int, *,
                         omml_element=None,
                         left: float = 1.0, top: float = 2.0,
                         width: float = 8.0, height: float = 1.0,
                         justification: str = "center",
                         name: str | None = None) -> str:
    """Add a display (standalone) equation to a slide.

    Creates a ``<m:oMathPara>`` with center justification.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    """
    from lxml import etree

    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    if slide_index < 1 or slide_index > len(prs.slides):
        raise IndexError(f"slide_index {slide_index} out of range (1..{len(prs.slides)})")
    try:
        slide = prs.slides[slide_index - 1]
        _name = name or f"DisplayEq {len(slide.shapes)}"

        txBox = slide.shapes.add_textbox(left, top, width, height)
        txBox.name = _name
        tf = txBox.text_frame
        tf.clear()

        p_elem = tf.paragraphs[0]._p

        if omml_element is not None:
            # Wrap in oMathPara for display mode
            oMathPara = etree.Element(f"{{{_NS_M}}}oMathPara")
            oMathParaPr = etree.SubElement(oMathPara, f"{{{_NS_M}}}oMathParaPr")
            jc = etree.SubElement(oMathParaPr, f"{{{_NS_M}}}jc")
            jc.set(f"{{{_NS_M}}}val", justification)
            oMathPara.append(omml_element)
            p_elem.append(oMathPara)

        return _name
    finally:
        _save_prs(prs, path)


# ---------------------------------------------------------------------------
# LaTeX to OMML converter (basic)
# ---------------------------------------------------------------------------

def latex_to_omml(latex: str) -> Any:
    """Convert a LaTeX math expression to an OMML element.

    Supports a subset of LaTeX math:
    - Fractions: ``\\frac{a}{b}``
    - Superscripts: ``x^{2}`` or ``x^2``
    - Subscripts: ``x_{i}`` or ``x_i``
    - Square roots: ``\\sqrt{x}`` or ``\\sqrt[n]{x}``
    - Sums/Integrals: ``\\sum_{i=0}^{n}``, ``\\int_{a}^{b}``
    - Delimiters: ``\\left(`` ... ``\\right)``
    - Greek letters: ``\\alpha``, ``\\beta``, etc.
    - Functions: ``\\sin``, ``\\cos``, ``\\log``, etc.

    Returns an ``<m:oMath>`` element.
    """
    builder = OmmlBuilder()
    _parse_latex_segment(latex, builder)
    return builder.to_xml()


# Greek letter mapping
_GREEK_MAP = {
    "\\alpha": "α", "\\beta": "β", "\\gamma": "γ", "\\delta": "δ",
    "\\epsilon": "ε", "\\zeta": "ζ", "\\eta": "η", "\\theta": "θ",
    "\\iota": "ι", "\\kappa": "κ", "\\lambda": "λ", "\\mu": "μ",
    "\\nu": "ν", "\\xi": "ξ", "\\pi": "π", "\\rho": "ρ",
    "\\sigma": "σ", "\\tau": "τ", "\\upsilon": "υ", "\\phi": "φ",
    "\\chi": "χ", "\\psi": "ψ", "\\omega": "ω",
    "\\Gamma": "Γ", "\\Delta": "Δ", "\\Theta": "Θ", "\\Lambda": "Λ",
    "\\Xi": "Ξ", "\\Pi": "Π", "\\Sigma": "Σ", "\\Phi": "Φ",
    "\\Psi": "Ψ", "\\Omega": "Ω",
    "\\nabla": "∇", "\\infty": "∞", "\\partial": "∂",
    "\\pm": "±", "\\mp": "∓", "\\times": "×", "\\div": "÷",
    "\\cdot": "·", "\\neq": "≠", "\\leq": "≤", "\\geq": "≥",
    "\\lt": "<", "\\gt": ">", "\\approx": "≈", "\\equiv": "≡",
    "\\in": "∈", "\\notin": "∉", "\\subset": "⊂", "\\supset": "⊃",
    "\\cup": "∪", "\\cap": "∩", "\\emptyset": "∅", "\\forall": "∀",
    "\\exists": "∃", "\\rightarrow": "→", "\\leftarrow": "←",
    "\\Rightarrow": "⇒", "\\Leftarrow": "⇐",
}

_FUNC_MAP = {
    "\\sin": "sin", "\\cos": "cos", "\\tan": "tan",
    "\\log": "log", "\\ln": "ln", "\\exp": "exp",
    "\\lim": "lim", "\\max": "max", "\\min": "min",
    "\\sup": "sup", "\\inf": "inf",
}


def _parse_latex_segment(latex: str, builder: OmmlBuilder) -> None:
    """Parse a LaTeX segment and add to the builder."""
    i = 0
    while i < len(latex):
        c = latex[i]

        # Skip whitespace
        if c in " \t\n":
            i += 1
            continue

        # Backslash commands
        if c == "\\":
            # Find the command
            j = i + 1
            while j < len(latex) and latex[j].isalpha():
                j += 1
            cmd = latex[i:j]

            if cmd == "\\frac":
                # Parse {num}{den}
                num, end1 = _extract_brace_content(latex, j)
                den, end2 = _extract_brace_content(latex, end1)
                sub_builder = OmmlBuilder()
                _parse_latex_segment(num, sub_builder)
                num_text = _builder_to_plain_text(sub_builder)
                sub_builder2 = OmmlBuilder()
                _parse_latex_segment(den, sub_builder2)
                den_text = _builder_to_plain_text(sub_builder2)
                builder.add_fraction(num_text, den_text)
                i = end2

            elif cmd == "\\sqrt":
                # Check for optional [degree]
                if j < len(latex) and latex[j] == "[":
                    bracket_end = latex.index("]", j)
                    degree = latex[j + 1:bracket_end]
                    j = bracket_end + 1
                else:
                    degree = ""
                base, end = _extract_brace_content(latex, j)
                if not base:
                    base = latex[j] if j < len(latex) else ""
                    end = j + 1
                builder.add_radical(base, degree)
                i = end

            elif cmd in ("\\sum", "\\prod", "\\coprod", "\\int", "\\iint",
                         "\\iiint", "\\oint", "\\bigcup", "\\bigcap",
                         "\\bigoplus", "\\bigotimes"):
                char_map = {
                    "\\sum": "∑", "\\prod": "∏", "\\coprod": "∐",
                    "\\int": "∫", "\\iint": "∬", "\\iiint": "∭",
                    "\\oint": "∮", "\\bigcup": "⋃", "\\bigcap": "⋂",
                    "\\bigoplus": "⨁", "\\bigotimes": "⨂",
                }
                nary_char = char_map.get(cmd, "∑")
                lower = ""
                upper = ""

                # Parse _{lower}^{upper} or _lower^upper
                nj = j
                if nj < len(latex) and latex[nj] == "_":
                    nj += 1
                    if nj < len(latex) and latex[nj] == "{":
                        lower, nj = _extract_brace_content(latex, nj)
                    else:
                        # Single char subscript
                        lower = latex[nj] if nj < len(latex) else ""
                        nj += 1

                if nj < len(latex) and latex[nj] == "^":
                    nj += 1
                    if nj < len(latex) and latex[nj] == "{":
                        upper, nj = _extract_brace_content(latex, nj)
                    else:
                        upper = latex[nj] if nj < len(latex) else ""
                        nj += 1

                # Rest is the base
                base = latex[nj:] if nj < len(latex) else ""
                builder.add_nary(base, lower, upper, char=nary_char)
                i = len(latex)  # Consumed rest

            elif cmd == "\\left":
                # Skip \left and the delimiter
                i = j + 1 if j < len(latex) else j
            elif cmd == "\\right":
                i = j + 1 if j < len(latex) else j

            elif cmd in _GREEK_MAP:
                builder.add_text_run(_GREEK_MAP[cmd])
                i = j

            elif cmd in _FUNC_MAP:
                # Functions render as text
                builder.add_text_run(_FUNC_MAP[cmd])
                i = j

            else:
                # Unknown command, render as text
                builder.add_text_run(cmd.replace("\\", ""))
                i = j
            continue

        # Superscript ^
        if c == "^":
            i += 1
            if i < len(latex) and latex[i] == "{":
                sup_content, end = _extract_brace_content(latex, i)
                builder.add_superscript("", sup_content)
                i = end
            elif i < len(latex):
                builder.add_superscript("", latex[i])
                i += 1
            continue

        # Subscript _
        if c == "_":
            i += 1
            if i < len(latex) and latex[i] == "{":
                sub_content, end = _extract_brace_content(latex, i)
                builder.add_subscript("", sub_content)
                i = end
            elif i < len(latex):
                builder.add_subscript("", latex[i])
                i += 1
            continue

        # Braces (just content grouping)
        if c == "{":
            content, end = _extract_brace_content(latex, i)
            _parse_latex_segment(content, builder)
            i = end
            continue

        # Regular character
        builder.add_text_run(c)
        i += 1


def _extract_brace_content(text: str, start: int) -> tuple[str, int]:
    """Extract content between balanced braces starting at position start.

    Returns (content, end_position) where end_position is after the closing brace.
    """
    if start >= len(text) or text[start] != "{":
        return "", start

    depth = 0
    i = start
    while i < len(text):
        if text[i] == "{":
            depth += 1
        elif text[i] == "}":
            depth -= 1
            if depth == 0:
                return text[start + 1:i], i + 1
        i += 1

    # Unbalanced braces
    return text[start + 1:], len(text)


def _builder_to_plain_text(builder: OmmlBuilder) -> str:
    """Extract plain text from a builder (simplified)."""
    from lxml import etree
    parts = []
    for t_elem in builder.to_xml().iter(f"{{{_NS_M}}}t"):
        if t_elem.text:
            parts.append(t_elem.text)
    return "".join(parts)


# ---------------------------------------------------------------------------
# List / remove equations
# ---------------------------------------------------------------------------

def list_equations(prs_or_path, slide_index: int) -> list[dict]:
    """List all equations on a slide.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).

    Returns a list of dicts with keys: name, is_display, text_preview.
    """
    from lxml import etree

    prs = _open_prs(prs_or_path)
    if slide_index < 1 or slide_index > len(prs.slides):
        raise IndexError(f"slide_index {slide_index} out of range (1..{len(prs.slides)})")
    try:
        slide = prs.slides[slide_index - 1]
        results = []

        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for para in shape.text_frame.paragraphs:
                p_elem = para._p
                # Check for inline equations
                for omath in p_elem.findall(f".//{{{_NS_M}}}oMath"):
                    text_parts = []
                    for t in omath.iter(f"{{{_NS_M}}}t"):
                        if t.text:
                            text_parts.append(t.text)
                    results.append({
                        "name": shape.name,
                        "is_display": False,
                        "text_preview": " ".join(text_parts)[:100],
                    })

                # Check for display equations
                for omath_para in p_elem.findall(f".//{{{_NS_M}}}oMathPara"):
                    text_parts = []
                    for t in omath_para.iter(f"{{{_NS_M}}}t"):
                        if t.text:
                            text_parts.append(t.text)
                    results.append({
                        "name": shape.name,
                        "is_display": True,
                        "text_preview": " ".join(text_parts)[:100],
                    })

        return results
    finally:
        pass


def remove_equation(prs_or_path, slide_index: int, shape_name: str) -> bool:
    """Remove all equations from a specific shape.

    This removes the OMML elements but keeps the text box.

    Parameters
    ----------
    slide_index : int
        1-based slide index (1 = first slide).
    """
    from lxml import etree

    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    if slide_index < 1 or slide_index > len(prs.slides):
        raise IndexError(f"slide_index {slide_index} out of range (1..{len(prs.slides)})")
    try:
        slide = prs.slides[slide_index - 1]
        shape = _find_shape(slide, shape_name)
        if shape is None:
            return False

        found = False
        if shape.has_text_frame:
            for para in shape.text_frame.paragraphs:
                p_elem = para._p
                for omath in p_elem.findall(f"{{{_NS_M}}}oMath"):
                    p_elem.remove(omath)
                    found = True
                for omath_para in p_elem.findall(f"{{{_NS_M}}}oMathPara"):
                    p_elem.remove(omath_para)
                    found = True

        return found
    finally:
        _save_prs(prs, path)


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _find_shape(slide, shape_name: str):
    for shape in slide.shapes:
        if shape.name == shape_name:
            return shape
    return None


def _is_presentation(obj) -> bool:
    """Check whether *obj* is a ``Presentation`` instance without eager import."""
    return type(obj).__name__ == "Presentation" and type(obj).__module__.startswith("pptx")


def _open_prs(prs_or_path):
    """Open a Presentation from *prs_or_path*.

    Accepts either an already-opened ``Presentation`` object or a file path.
    Returns the ``Presentation`` object directly.
    """
    from pptx import Presentation

    if _is_presentation(prs_or_path):
        return prs_or_path
    return Presentation(str(prs_or_path))


def _save_prs(prs, path):
    """Save *prs* back to *path* if *path* is not None."""
    if path is not None:
        prs.save(str(path))
