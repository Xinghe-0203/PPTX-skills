"""SVG vector graphics to native PowerPoint shapes (DrawingML) conversion.

Converts SVG ``<path>``, ``<rect>``, ``<circle>``, ``<ellipse>``, ``<line>``,
``<polygon>``, and ``<polyline>`` elements into ``<a:custGeom>`` (custom
geometry) shapes that are fully editable in PowerPoint.

For SVGs containing features that cannot be represented in DrawingML (filters,
gradients, text on path, clip-paths, masks, etc.), a fallback rasterisation
path renders the SVG to PNG via Pillow/Cairo and embeds it as a picture.

OOXML reference: ECMA-376 Part 4, §20.1.9 (Custom Geometry),
§20.1.2 (Line Properties), §20.1.4 (Fill Properties).
"""

from __future__ import annotations

import math
import re
from dataclasses import dataclass, field
from typing import Any

__all__ = [
    "PathData",
    "import_svg",
    "import_svg_as_image",
    "svg_to_drawingml",
    "list_svg_shapes",
]

# ---------------------------------------------------------------------------
# Namespace constants
# ---------------------------------------------------------------------------

_NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
_NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# ---------------------------------------------------------------------------
# Unit helpers
# ---------------------------------------------------------------------------

_PT_TO_EMU = 12700
_INCH_TO_EMU = 914400


def _pt_to_emu(pt: float) -> int:
    return int(round(pt * _PT_TO_EMU))


def _inch_to_emu(inch: float) -> int:
    return int(round(inch * _INCH_TO_EMU))


# ---------------------------------------------------------------------------
# SVG named colours (CSS4 subset)
# ---------------------------------------------------------------------------

_SVG_NAMED_COLORS: dict[str, str] = {
    "black": "000000", "white": "FFFFFF", "red": "FF0000",
    "green": "008000", "blue": "0000FF", "yellow": "FFFF00",
    "cyan": "00FFFF", "magenta": "FF00FF", "silver": "C0C0C0",
    "gray": "808080", "grey": "808080", "maroon": "800000",
    "olive": "808000", "lime": "00FF00", "aqua": "00FFFF",
    "teal": "008080", "navy": "000080", "fuchsia": "FF00FF",
    "purple": "800080", "orange": "FFA500", "aliceblue": "F0F8FF",
    "antiquewhite": "FAEBD7", "aquamarine": "7FFFD4",
    "azure": "F0FFFF", "beige": "F5F5DC", "bisque": "FFE4C4",
    "blanchedalmond": "FFEBCD", "blueviolet": "8A2BE2",
    "brown": "A52A2A", "burlywood": "DEB887", "cadetblue": "5F9EA0",
    "chartreuse": "7FFF00", "chocolate": "D2691E",
    "coral": "FF7F50", "cornflowerblue": "6495ED",
    "cornsilk": "FFF8DC", "crimson": "DC143C", "darkblue": "00008B",
    "darkcyan": "008B8B", "darkgoldenrod": "B8860B",
    "darkgray": "A9A9A9", "darkgrey": "A9A9A9",
    "darkgreen": "006400", "darkkhaki": "BDB76B",
    "darkmagenta": "8B008B", "darkolivegreen": "556B2F",
    "darkorange": "FF8C00", "darkorchid": "9932CC",
    "darkred": "8B0000", "darksalmon": "E9967A",
    "darkseagreen": "8FBC8F", "darkslateblue": "483D8B",
    "darkslategray": "2F4F4F", "darkslategrey": "2F4F4F",
    "darkturquoise": "00CED1", "darkviolet": "9400D3",
    "deeppink": "FF1493", "deepskyblue": "00BFFF",
    "dimgray": "696969", "dimgrey": "696969",
    "dodgerblue": "1E90FF", "firebrick": "B22222",
    "floralwhite": "FFFAF0", "forestgreen": "228B22",
    "gainsboro": "DCDCDC", "ghostwhite": "F8F8FF",
    "gold": "FFD700", "goldenrod": "DAA520", "greenyellow": "ADFF2F",
    "honeydew": "F0FFF0", "hotpink": "FF69B4",
    "indianred": "CD5C5C", "indigo": "4B0082", "ivory": "FFFFF0",
    "khaki": "F0E68C", "lavender": "E6E6FA", "lavenderblush": "FFF0F5",
    "lawngreen": "7CFC00", "lemonchiffon": "FFFACD",
    "lightblue": "ADD8E6", "lightcoral": "F08080",
    "lightcyan": "E0FFFF", "lightgoldenrodyellow": "FAFAD2",
    "lightgray": "D3D3D3", "lightgrey": "D3D3D3",
    "lightgreen": "90EE90", "lightpink": "FFB6C1",
    "lightsalmon": "FFA07A", "lightseagreen": "20B2AA",
    "lightskyblue": "87CEFA", "lightslategray": "778899",
    "lightslategrey": "778899", "lightsteelblue": "B0C4DE",
    "lightyellow": "FFFFE0", "limegreen": "32CD32",
    "linen": "FAF0E6", "mediumaquamarine": "66CDAA",
    "mediumblue": "0000CD", "mediumorchid": "BA55D3",
    "mediumpurple": "9370DB", "mediumseagreen": "3CB371",
    "mediumslateblue": "7B68EE", "mediumspringgreen": "00FA9A",
    "mediumturquoise": "48D1CC", "mediumvioletred": "C71585",
    "midnightblue": "191970", "mintcream": "F5FFFA",
    "mistyrose": "FFE4E1", "moccasin": "FFE4B5",
    "navajowhite": "FFDEAD", "oldlace": "FDF5E6",
    "olivedrab": "6B8E23", "orangered": "FF4500",
    "orchid": "DA70D6", "palegoldenrod": "EEE8AA",
    "palegreen": "98FB98", "paleturquoise": "AFEEEE",
    "palevioletred": "DB7093", "papayawhip": "FFEFD5",
    "peachpuff": "FFDAB9", "peru": "CD853F", "pink": "FFC0CB",
    "plum": "DDA0DD", "powderblue": "B0E0E6",
    "rosybrown": "BC8F8F", "royalblue": "4169E1",
    "saddlebrown": "8B4513", "salmon": "FA8072",
    "sandybrown": "F4A460", "seagreen": "2E8B57",
    "seashell": "FFF5EE", "sienna": "A0522D",
    "skyblue": "87CEEB", "slateblue": "6A5ACD",
    "slategray": "708090", "slategrey": "708090",
    "snow": "FFFAFA", "springgreen": "00FF7F",
    "steelblue": "4682B4", "tan": "D2B48C", "thistle": "D8BFD8",
    "tomato": "FF6347", "turquoise": "40E0D0",
    "violet": "EE82EE", "wheat": "F5DEB3", "whitesmoke": "F5F5F5",
    "yellowgreen": "9ACD32",
}

# ---------------------------------------------------------------------------
# Unsupported SVG feature markers
# ---------------------------------------------------------------------------

_UNSUPPORTED_SVG_ELEMENTS = frozenset({
    "filter", "clippath", "mask", "symbol", "use", "image",
    "text", "tspan", "textpath",
})

_UNSUPPORTED_STYLE_KEYWORDS = frozenset({
    "filter", "clip-path", "mask", "gradient", "pattern",
    "text-path", "marker",
})


# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------

@dataclass
class PathData:
    """DrawingML path data for a single SVG element.

    Attributes
    ----------
    d : str
        DrawingML path commands string (moveTo, lnTo, cubicBezTo, close).
    fill : str | None
        Fill colour as hex string (e.g. ``"FF0000"``), or ``None`` for no fill.
    stroke : str | None
        Stroke colour as hex string, or ``None`` for no stroke.
    stroke_width : float
        Stroke width in points.
    name : str
        Shape name identifier.
    """

    d: str
    fill: str | None = None
    stroke: str | None = None
    stroke_width: float = 1.0
    name: str = ""


# ---------------------------------------------------------------------------
# SVG path data parser
# ---------------------------------------------------------------------------

def _tokenize_path(d: str) -> list[str]:
    """Tokenize an SVG path ``d`` attribute into commands and numbers.

    Handles multi-digit numbers, negative numbers, decimal points,
    scientific notation, and comma/whitespace separators.
    """
    tokens: list[str] = []
    i = 0
    n = len(d)
    while i < n:
        c = d[i]
        # Skip whitespace and commas
        if c in " \t\n\r,":
            i += 1
            continue
        # Command letter
        if c.isalpha():
            tokens.append(c)
            i += 1
            continue
        # Number: optional sign, digits, optional decimal, optional exponent
        start = i
        if c in "+-":
            i += 1
        has_dot = False
        has_exp = False
        while i < n:
            ch = d[i]
            if ch.isdigit():
                i += 1
            elif ch == "." and not has_dot and not has_exp:
                has_dot = True
                i += 1
            elif ch in "eE" and not has_exp:
                has_exp = True
                i += 1
                if i < n and d[i] in "+-":
                    i += 1
            else:
                break
        if i > start:
            tokens.append(d[start:i])
        else:
            # Safety: skip unexpected character
            i += 1
    return tokens


def _parse_numbers(tokens: list[str], start: int, count: int) -> tuple[list[float], int]:
    """Parse *count* numbers from *tokens* starting at *start*.

    Returns (values, new_index).
    """
    values: list[float] = []
    idx = start
    for _ in range(count):
        if idx < len(tokens):
            try:
                values.append(float(tokens[idx]))
            except ValueError:
                break
            idx += 1
        else:
            break
    return values, idx


# ---------------------------------------------------------------------------
# Arc-to-cubic-bezier approximation
# ---------------------------------------------------------------------------

def _arc_to_cubics(
    x1: float, y1: float,
    rx: float, ry: float,
    phi: float,
    fA: bool, fS: bool,
    x2: float, y2: float,
) -> list[tuple[float, float, float, float, float, float]]:
    """Convert an SVG arc to a sequence of cubic Bezier curves.

    Uses the standard parametric conversion algorithm from the SVG spec
    (Appendix F).  Returns a list of ``(cp1x, cp1y, cp2x, cp2y, x, y)``
    tuples, one per cubic segment (4 segments per full semicircle).
    """
    if rx == 0 or ry == 0:
        return [((x1 + x2) / 2, (y1 + y2) / 2, (x1 + x2) / 2, (y1 + y2) / 2, x2, y2)]

    # Ensure radii are positive
    rx = abs(rx)
    ry = abs(ry)

    cos_phi = math.cos(phi)
    sin_phi = math.sin(phi)

    # Step 1: compute (x1', y1')
    dx = (x1 - x2) / 2
    dy = (y1 - y2) / 2
    x1p = cos_phi * dx + sin_phi * dy
    y1p = -sin_phi * dx + cos_phi * dy

    # Step 2: compute (cx', cy')
    x1p_sq = x1p * x1p
    y1p_sq = y1p * y1p
    rx_sq = rx * rx
    ry_sq = ry * ry

    # Correct out-of-range radii
    lam = x1p_sq / rx_sq + y1p_sq / ry_sq
    if lam > 1:
        lam_sqrt = math.sqrt(lam)
        rx *= lam_sqrt
        ry *= lam_sqrt
        rx_sq = rx * rx
        ry_sq = ry * ry

    num = rx_sq * ry_sq - rx_sq * y1p_sq - ry_sq * x1p_sq
    den = rx_sq * y1p_sq + ry_sq * x1p_sq

    if den == 0:
        return [((x1 + x2) / 2, (y1 + y2) / 2, (x1 + x2) / 2, (y1 + y2) / 2, x2, y2)]

    sq = max(0.0, num / den)
    sq = math.sqrt(sq)
    if fA == fS:
        sq = -sq

    cxp = sq * rx * y1p / ry
    cyp = -sq * ry * x1p / rx

    # Step 3: compute (cx, cy)
    cx = cos_phi * cxp - sin_phi * cyp + (x1 + x2) / 2
    cy = sin_phi * cxp + cos_phi * cyp + (y1 + y2) / 2

    # Step 4: compute theta1 and dtheta
    def _angle(ux: float, uy: float, vx: float, vy: float) -> float:
        dot = ux * vx + uy * vy
        len_u = math.sqrt(ux * ux + uy * uy)
        len_v = math.sqrt(vx * vx + vy * vy)
        if len_u == 0 or len_v == 0:
            return 0.0
        cos_a = max(-1.0, min(1.0, dot / (len_u * len_v)))
        a = math.acos(cos_a)
        if ux * vy - uy * vx < 0:
            a = -a
        return a

    theta1 = _angle(1, 0, (x1p - cxp) / rx, (y1p - cyp) / ry)
    dtheta = _angle(
        (x1p - cxp) / rx, (y1p - cyp) / ry,
        (-x1p - cxp) / rx, (-y1p - cyp) / ry,
    )

    if fS and dtheta < 0:
        dtheta += 2 * math.pi
    elif not fS and dtheta > 0:
        dtheta -= 2 * math.pi

    # Split into segments (approximation uses 4 segments per semicircle)
    n_segs = max(1, int(math.ceil(abs(dtheta) / (math.pi / 2))))
    step = dtheta / n_segs

    cubics: list[tuple[float, float, float, float, float, float]] = []
    for i in range(n_segs):
        t1 = theta1 + step * i
        t2 = theta1 + step * (i + 1)
        alpha = step / 2
        # Tangent at t1
        cos_t1 = math.cos(t1)
        sin_t1 = math.sin(t1)
        cos_t2 = math.cos(t2)
        sin_t2 = math.sin(t2)

        # Endpoint
        ex = cos_phi * rx * cos_t2 - sin_phi * ry * sin_t2 + cx
        ey = sin_phi * rx * cos_t2 + cos_phi * ry * sin_t2 + cy

        # Control point 1
        k1 = (4.0 / 3.0) * math.tan(alpha)
        cp1x = cos_phi * rx * (cos_t1 - k1 * sin_t1) - sin_phi * ry * (sin_t1 + k1 * cos_t1) + cx
        cp1y = sin_phi * rx * (cos_t1 - k1 * sin_t1) + cos_phi * ry * (sin_t1 + k1 * cos_t1) + cy

        # Control point 2
        k2 = (4.0 / 3.0) * math.tan(alpha)
        cp2x = cos_phi * rx * (cos_t2 + k2 * sin_t2) - sin_phi * ry * (sin_t2 - k2 * cos_t2) + cx
        cp2y = sin_phi * rx * (cos_t2 + k2 * sin_t2) + cos_phi * ry * (sin_t2 - k2 * cos_t2) + cy

        cubics.append((cp1x, cp1y, cp2x, cp2y, ex, ey))

    return cubics


# ---------------------------------------------------------------------------
# SVG path command interpreter
# ---------------------------------------------------------------------------

class _PathInterpreter:
    """Interpret SVG path *d* tokens and produce DrawingML path commands.

    DrawingML commands emitted:
      - ``moveTo x y``
      - ``lnTo x y``
      - ``cubicBezTo cp1x cp1y cp2x cp2y x y``
      - ``close``
    """

    def __init__(self) -> None:
        self.commands: list[str] = []
        self._sx: float = 0.0  # start of current sub-path
        self._sy: float = 0.0
        self._cx: float = 0.0  # current point
        self._cy: float = 0.0
        self._last_cpx: float = 0.0  # last control point (for S/s/T/t)
        self._last_cpy: float = 0.0
        self._prev_cmd: str = ""

    def interpret(self, d: str) -> list[str]:
        """Parse and interpret an SVG path *d* attribute string."""
        tokens = _tokenize_path(d)
        i = 0
        n = len(tokens)

        while i < n:
            cmd = tokens[i]
            if cmd.isalpha():
                i += 1
            else:
                # Implicit repeat of previous command
                cmd = self._prev_cmd
                if cmd == "":
                    break

            cmd_upper = cmd.upper()

            if cmd_upper == "M":
                i = self._do_moveto(cmd, tokens, i)
            elif cmd_upper == "L":
                i = self._do_lineto(cmd, tokens, i)
            elif cmd_upper == "H":
                i = self._do_hlineto(cmd, tokens, i)
            elif cmd_upper == "V":
                i = self._do_vlineto(cmd, tokens, i)
            elif cmd_upper == "C":
                i = self._do_curveto(cmd, tokens, i)
            elif cmd_upper == "S":
                i = self._do_smooth_curveto(cmd, tokens, i)
            elif cmd_upper == "Q":
                i = self._do_quadto(cmd, tokens, i)
            elif cmd_upper == "T":
                i = self._do_smooth_quadto(cmd, tokens, i)
            elif cmd_upper == "A":
                i = self._do_arc(cmd, tokens, i)
            elif cmd_upper == "Z":
                self._do_close()
            else:
                # Unknown command — skip
                i += 1
                continue

            self._prev_cmd = cmd

        return self.commands

    # -- Moveto -------------------------------------------------------------

    def _do_moveto(self, cmd: str, tokens: list[str], i: int) -> int:
        vals, i = _parse_numbers(tokens, i, 2)
        if len(vals) < 2:
            return i
        if cmd == "M":
            self._cx, self._cy = vals[0], vals[1]
        else:  # "m"
            self._cx += vals[0]
            self._cy += vals[1]
        self._sx = self._cx
        self._sy = self._cy
        self.commands.append(f"moveTo {self._cx} {self._cy}")
        # Implicit lineto for subsequent coordinate pairs
        while i + 1 < len(tokens) and not tokens[i].isalpha():
            vals, i = _parse_numbers(tokens, i, 2)
            if len(vals) < 2:
                break
            if cmd == "M":
                self._cx, self._cy = vals[0], vals[1]
            else:
                self._cx += vals[0]
                self._cy += vals[1]
            self.commands.append(f"lnTo {self._cx} {self._cy}")
        return i

    # -- Lineto -------------------------------------------------------------

    def _do_lineto(self, cmd: str, tokens: list[str], i: int) -> int:
        while i + 1 < len(tokens) and not tokens[i].isalpha():
            vals, i = _parse_numbers(tokens, i, 2)
            if len(vals) < 2:
                break
            if cmd == "L":
                self._cx, self._cy = vals[0], vals[1]
            else:
                self._cx += vals[0]
                self._cy += vals[1]
            self.commands.append(f"lnTo {self._cx} {self._cy}")
        return i

    # -- Horizontal lineto --------------------------------------------------

    def _do_hlineto(self, cmd: str, tokens: list[str], i: int) -> int:
        while i < len(tokens) and not tokens[i].isalpha():
            vals, i = _parse_numbers(tokens, i, 1)
            if len(vals) < 1:
                break
            if cmd == "H":
                self._cx = vals[0]
            else:
                self._cx += vals[0]
            self.commands.append(f"lnTo {self._cx} {self._cy}")
        return i

    # -- Vertical lineto ----------------------------------------------------

    def _do_vlineto(self, cmd: str, tokens: list[str], i: int) -> int:
        while i < len(tokens) and not tokens[i].isalpha():
            vals, i = _parse_numbers(tokens, i, 1)
            if len(vals) < 1:
                break
            if cmd == "V":
                self._cy = vals[0]
            else:
                self._cy += vals[0]
            self.commands.append(f"lnTo {self._cx} {self._cy}")
        return i

    # -- Cubic Bezier -------------------------------------------------------

    def _do_curveto(self, cmd: str, tokens: list[str], i: int) -> int:
        while i + 5 < len(tokens) and not tokens[i].isalpha():
            vals, i = _parse_numbers(tokens, i, 6)
            if len(vals) < 6:
                break
            if cmd == "C":
                cp1x, cp1y, cp2x, cp2y, ex, ey = vals
            else:
                cp1x = self._cx + vals[0]
                cp1y = self._cy + vals[1]
                cp2x = self._cx + vals[2]
                cp2y = self._cy + vals[3]
                ex = self._cx + vals[4]
                ey = self._cy + vals[5]
            self._last_cpx = cp2x
            self._last_cpy = cp2y
            self._cx = ex
            self._cy = ey
            self.commands.append(
                f"cubicBezTo {cp1x} {cp1y} {cp2x} {cp2y} {ex} {ey}"
            )
        return i

    # -- Smooth cubic Bezier ------------------------------------------------

    def _do_smooth_curveto(self, cmd: str, tokens: list[str], i: int) -> int:
        while i + 3 < len(tokens) and not tokens[i].isalpha():
            vals, i = _parse_numbers(tokens, i, 4)
            if len(vals) < 4:
                break
            # Compute reflected control point
            if self._prev_cmd.upper() in ("C", "S"):
                cp1x = 2 * self._cx - self._last_cpx
                cp1y = 2 * self._cy - self._last_cpy
            else:
                cp1x = self._cx
                cp1y = self._cy

            if cmd == "S":
                cp2x, cp2y, ex, ey = vals
            else:
                cp2x = self._cx + vals[0]
                cp2y = self._cy + vals[1]
                ex = self._cx + vals[2]
                ey = self._cy + vals[3]

            self._last_cpx = cp2x
            self._last_cpy = cp2y
            self._cx = ex
            self._cy = ey
            self.commands.append(
                f"cubicBezTo {cp1x} {cp1y} {cp2x} {cp2y} {ex} {ey}"
            )
        return i

    # -- Quadratic Bezier ---------------------------------------------------

    def _do_quadto(self, cmd: str, tokens: list[str], i: int) -> int:
        while i + 3 < len(tokens) and not tokens[i].isalpha():
            vals, i = _parse_numbers(tokens, i, 4)
            if len(vals) < 4:
                break
            if cmd == "Q":
                qpx, qpy, ex, ey = vals
            else:
                qpx = self._cx + vals[0]
                qpy = self._cy + vals[1]
                ex = self._cx + vals[2]
                ey = self._cy + vals[3]

            # Quadratic -> cubic conversion
            cp1x = self._cx + 2 * (qpx - self._cx) / 3
            cp1y = self._cy + 2 * (qpy - self._cy) / 3
            cp2x = ex + 2 * (qpx - ex) / 3
            cp2y = ey + 2 * (qpy - ey) / 3

            self._last_cpx = qpx
            self._last_cpy = qpy
            self._cx = ex
            self._cy = ey
            self.commands.append(
                f"cubicBezTo {cp1x} {cp1y} {cp2x} {cp2y} {ex} {ey}"
            )
        return i

    # -- Smooth quadratic Bezier --------------------------------------------

    def _do_smooth_quadto(self, cmd: str, tokens: list[str], i: int) -> int:
        while i + 1 < len(tokens) and not tokens[i].isalpha():
            vals, i = _parse_numbers(tokens, i, 2)
            if len(vals) < 2:
                break
            if cmd == "T":
                ex, ey = vals
            else:
                ex = self._cx + vals[0]
                ey = self._cy + vals[1]

            # Reflected control point
            if self._prev_cmd.upper() in ("Q", "T"):
                qpx = 2 * self._cx - self._last_cpx
                qpy = 2 * self._cy - self._last_cpy
            else:
                qpx = self._cx
                qpy = self._cy

            # Quadratic -> cubic conversion
            cp1x = self._cx + 2 * (qpx - self._cx) / 3
            cp1y = self._cy + 2 * (qpy - self._cy) / 3
            cp2x = ex + 2 * (qpx - ex) / 3
            cp2y = ey + 2 * (qpy - ey) / 3

            self._last_cpx = qpx
            self._last_cpy = qpy
            self._cx = ex
            self._cy = ey
            self.commands.append(
                f"cubicBezTo {cp1x} {cp1y} {cp2x} {cp2y} {ex} {ey}"
            )
        return i

    # -- Arc ----------------------------------------------------------------

    def _do_arc(self, cmd: str, tokens: list[str], i: int) -> int:
        while i + 6 < len(tokens) and not tokens[i].isalpha():
            vals, i = _parse_numbers(tokens, i, 7)
            if len(vals) < 7:
                break
            rx, ry = vals[0], vals[1]
            x_axis_rot = vals[2]
            large_arc = vals[3] != 0
            sweep = vals[4] != 0

            if cmd == "A":
                ex, ey = vals[5], vals[6]
            else:
                ex = self._cx + vals[5]
                ey = self._cy + vals[6]

            phi = math.radians(x_axis_rot)
            cubics = _arc_to_cubics(
                self._cx, self._cy, rx, ry, phi, large_arc, sweep, ex, ey,
            )
            for cp1x, cp1y, cp2x, cp2y, endx, endy in cubics:
                self.commands.append(
                    f"cubicBezTo {cp1x} {cp1y} {cp2x} {cp2y} {endx} {endy}"
                )

            self._cx = ex
            self._cy = ey
            # Reset last control point after arc
            self._last_cpx = self._cx
            self._last_cpy = self._cy
        return i

    # -- Close --------------------------------------------------------------

    def _do_close(self) -> None:
        self.commands.append("close")
        self._cx = self._sx
        self._cy = self._sy


# ---------------------------------------------------------------------------
# SVG colour parsing
# ---------------------------------------------------------------------------

def _parse_svg_color(color_str: str | None) -> str | None:
    """Parse an SVG colour value to a hex string (without ``#``).

    Supports ``#RGB``, ``#RRGGBB``, ``rgb(r,g,b)``, ``rgba(r,g,b,a)``,
    and CSS named colours.

    Returns ``None`` for ``"none"`` / empty / unparsable values.
    """
    if not color_str or color_str.strip().lower() == "none":
        return None

    s = color_str.strip()

    # Hex: #RGB or #RRGGBB
    if s.startswith("#"):
        hex_part = s[1:]
        if len(hex_part) == 3:
            return (hex_part[0] * 2 + hex_part[1] * 2 + hex_part[2] * 2).upper()
        if len(hex_part) == 6:
            return hex_part.upper()
        return None

    # rgb() / rgba() — handles both absolute and percentage values
    m = re.match(r"rgba?\(\s*([^)]+)\)", s)
    if m:
        parts = [p.strip().rstrip("%") for p in m.group(1).split(",")]
        if len(parts) < 3:
            return None
        # Detect whether values are percentages
        raw_parts = [p.strip() for p in m.group(1).split(",")]
        is_pct = "%" in raw_parts[0]
        try:
            if is_pct:
                r = int(round(float(parts[0]) * 2.55))
                g = int(round(float(parts[1]) * 2.55))
                b = int(round(float(parts[2]) * 2.55))
            else:
                r = int(float(parts[0]))
                g = int(float(parts[1]))
                b = int(float(parts[2]))
            r = max(0, min(255, r))
            g = max(0, min(255, g))
            b = max(0, min(255, b))
            return f"{r:02X}{g:02X}{b:02X}"
        except (ValueError, IndexError):
            return None

    # Named colour
    lower = s.lower()
    if lower in _SVG_NAMED_COLORS:
        return _SVG_NAMED_COLORS[lower]

    return None


def _parse_svg_opacity(opacity_str: str | None) -> float:
    """Parse an SVG opacity value to a 0-1 float."""
    if not opacity_str:
        return 1.0
    try:
        return max(0.0, min(1.0, float(opacity_str)))
    except ValueError:
        return 1.0


# ---------------------------------------------------------------------------
# SVG transform parsing
# ---------------------------------------------------------------------------

def _parse_transform(transform_str: str | None) -> list[list[float]]:
    """Parse an SVG ``transform`` attribute into a list of matrix operations.

    Returns a list of ``[a, b, c, d, e, f]`` affine matrices, to be applied
    left-to-right (i.e. concatenated).
    """
    if not transform_str:
        return []

    matrices: list[list[float]] = []
    # Match transform functions
    for m in re.finditer(
        r"(matrix|translate|scale|rotate|skewX|skewY)\s*\(([^)]*)\)",
        transform_str,
    ):
        func = m.group(1)
        args = [float(x) for x in re.split(r"[\s,]+", m.group(2).strip()) if x]

        if func == "matrix" and len(args) == 6:
            matrices.append(args)
        elif func == "translate":
            tx = args[0] if len(args) >= 1 else 0
            ty = args[1] if len(args) >= 2 else 0
            matrices.append([1, 0, 0, 1, tx, ty])
        elif func == "scale":
            sx = args[0] if len(args) >= 1 else 1
            sy = args[1] if len(args) >= 2 else sx
            matrices.append([sx, 0, 0, sy, 0, 0])
        elif func == "rotate" and len(args) >= 1:
            angle = math.radians(args[0])
            cos_a = math.cos(angle)
            sin_a = math.sin(angle)
            if len(args) >= 3:
                cx, cy = args[1], args[2]
                # rotate(a, cx, cy) = translate(cx,cy) * rotate(a) * translate(-cx,-cy)
                matrices.append([1, 0, 0, 1, cx, cy])
                matrices.append([cos_a, sin_a, -sin_a, cos_a, 0, 0])
                matrices.append([1, 0, 0, 1, -cx, -cy])
            else:
                matrices.append([cos_a, sin_a, -sin_a, cos_a, 0, 0])
        elif func == "skewX" and len(args) >= 1:
            angle = math.radians(args[0])
            matrices.append([1, 0, math.tan(angle), 1, 0, 0])
        elif func == "skewY" and len(args) >= 1:
            angle = math.radians(args[0])
            matrices.append([1, math.tan(angle), 0, 1, 0, 0])

    return matrices


def _concat_matrices(matrices: list[list[float]]) -> list[float]:
    """Concatenate a list of affine matrices into one.

    Each matrix is ``[a, b, c, d, e, f]`` representing::

        | a c e |
        | b d f |
        | 0 0 1 |
    """
    if not matrices:
        return [1, 0, 0, 1, 0, 0]
    result = matrices[0]
    for mat in matrices[1:]:
        a = result[0] * mat[0] + result[2] * mat[1]
        b = result[1] * mat[0] + result[3] * mat[1]
        c = result[0] * mat[2] + result[2] * mat[3]
        d = result[1] * mat[2] + result[3] * mat[3]
        e = result[0] * mat[4] + result[2] * mat[5] + result[4]
        f = result[1] * mat[4] + result[3] * mat[5] + result[5]
        result = [a, b, c, d, e, f]
    return result


def _apply_matrix(matrix: list[float], x: float, y: float) -> tuple[float, float]:
    """Apply affine matrix ``[a, b, c, d, e, f]`` to point ``(x, y)``."""
    a, b, c, d, e, f = matrix
    nx = a * x + c * y + e
    ny = b * x + d * y + f
    return nx, ny


# ---------------------------------------------------------------------------
# SVG element → DrawingML path data
# ---------------------------------------------------------------------------

def _parse_svg_viewbox(root: Any) -> tuple[float, float, float, float]:
    """Extract viewBox or width/height from SVG root.

    Returns ``(x, y, width, height)``.
    """
    import xml.etree.ElementTree as ET

    # Try viewBox first
    vb = root.get("viewBox")
    if vb:
        parts = re.split(r"[\s,]+", vb.strip())
        if len(parts) >= 4:
            try:
                return float(parts[0]), float(parts[1]), float(parts[2]), float(parts[3])
            except ValueError:
                pass

    # Fallback to width/height
    w = root.get("width", "100")
    h = root.get("height", "100")
    # Strip units (px, pt, etc.)
    w = re.sub(r"[a-zA-Z]+$", "", w)
    h = re.sub(r"[a-zA-Z]+$", "", h)
    try:
        return 0.0, 0.0, float(w), float(h)
    except ValueError:
        return 0.0, 0.0, 100.0, 100.0


def _get_attr(element: Any, attr: str, default: str | None = None) -> str | None:
    """Get an attribute from an SVG element, checking style attribute first."""
    import xml.etree.ElementTree as ET

    # Check style attribute
    style = element.get("style", "")
    if style:
        for decl in style.split(";"):
            decl = decl.strip()
            if ":" in decl:
                k, v = decl.split(":", 1)
                if k.strip().lower() == attr.lower():
                    return v.strip()

    return element.get(attr, default)


def _element_has_unsupported_features(element: Any) -> bool:
    """Check if an SVG element uses features that cannot be converted to DrawingML."""
    import xml.etree.ElementTree as ET

    tag = element.tag
    # Strip namespace
    if "}" in tag:
        tag = tag.split("}", 1)[1]
    tag = tag.lower()

    if tag in _UNSUPPORTED_SVG_ELEMENTS:
        return True

    # Check style for unsupported properties
    style = element.get("style", "")
    fill = element.get("fill", "")
    stroke_val = element.get("stroke", "")

    for prop_source in [style, fill, stroke_val]:
        lower = prop_source.lower()
        if "url(" in lower:
            # Gradient or pattern reference
            return True
        for kw in _UNSUPPORTED_STYLE_KEYWORDS:
            if kw in lower:
                return True

    # Check specific gradient/pattern attributes
    if element.get("fill", "").startswith("url("):
        return True
    if element.get("stroke", "").startswith("url("):
        return True

    return False


def _svg_has_unsupported_features(root: Any) -> bool:
    """Recursively check if the SVG tree has any unsupported features."""
    import xml.etree.ElementTree as ET

    for elem in root.iter():
        if _element_has_unsupported_features(elem):
            return True
    return False


def _convert_path_element(
    elem: Any,
    transform_matrix: list[float],
    idx: int,
    parent_fill: str | None = None,
    parent_stroke: str | None = None,
    parent_stroke_width: float | None = None,
    parent_opacity: float = 1.0,
) -> PathData | None:
    """Convert an SVG ``<path>`` element to PathData.

    Applies *transform_matrix* to all coordinates in the parsed path commands.
    """
    d_attr = elem.get("d")
    if not d_attr:
        return None

    interpreter = _PathInterpreter()
    commands = interpreter.interpret(d_attr)
    if not commands:
        return None

    # Apply transform to coordinate values in commands
    transformed: list[str] = []
    for cmd_str in commands:
        parts = cmd_str.split()
        cmd_name = parts[0]
        if cmd_name == "close":
            transformed.append("close")
            continue

        coords = [float(x) for x in parts[1:]]
        transformed_parts = [cmd_name]
        if cmd_name == "moveTo" and len(coords) >= 2:
            x, y = _apply_matrix(transform_matrix, coords[0], coords[1])
            transformed_parts.extend([str(x), str(y)])
        elif cmd_name == "lnTo" and len(coords) >= 2:
            x, y = _apply_matrix(transform_matrix, coords[0], coords[1])
            transformed_parts.extend([str(x), str(y)])
        elif cmd_name == "cubicBezTo" and len(coords) >= 6:
            cp1x, cp1y = _apply_matrix(transform_matrix, coords[0], coords[1])
            cp2x, cp2y = _apply_matrix(transform_matrix, coords[2], coords[3])
            ex, ey = _apply_matrix(transform_matrix, coords[4], coords[5])
            transformed_parts.extend([str(cp1x), str(cp1y), str(cp2x), str(cp2y), str(ex), str(ey)])
        else:
            transformed.append(cmd_str)
            continue

        transformed.append(" ".join(transformed_parts))

    # Determine fill/stroke
    fill_str = _get_attr(elem, "fill")
    stroke_str = _get_attr(elem, "stroke")
    sw_str = _get_attr(elem, "stroke-width")
    opacity_str = _get_attr(elem, "opacity")
    fill_opacity_str = _get_attr(elem, "fill-opacity")
    stroke_opacity_str = _get_attr(elem, "stroke-opacity")

    # Inherit from parent if not explicitly set
    raw_fill = fill_str if fill_str is not None else parent_fill
    raw_stroke = stroke_str if stroke_str is not None else parent_stroke
    raw_sw = sw_str if sw_str is not None else parent_stroke_width

    fill_hex = _parse_svg_color(raw_fill)
    stroke_hex = _parse_svg_color(raw_stroke)

    # Handle opacity
    elem_opacity = _parse_svg_opacity(opacity_str)
    effective_opacity = parent_opacity * elem_opacity

    # Stroke width
    stroke_width = 1.0
    if raw_sw is not None:
        try:
            stroke_width = float(re.sub(r"[a-zA-Z]+$", "", str(raw_sw)))
        except ValueError:
            stroke_width = 1.0

    # SVG default fill is black — only apply if not explicitly "none"
    if fill_str is None and parent_fill is None and fill_hex is None:
        # SVG default: fill black
        fill_hex = "000000"

    name = elem.get("id", f"SVG_Path_{idx}")

    return PathData(
        d="\n".join(transformed),
        fill=fill_hex,
        stroke=stroke_hex,
        stroke_width=stroke_width * effective_opacity if stroke_hex else stroke_width,
        name=name,
    )


def _convert_rect_element(
    elem: Any,
    transform_matrix: list[float],
    idx: int,
    parent_fill: str | None = None,
    parent_stroke: str | None = None,
    parent_stroke_width: float | None = None,
    parent_opacity: float = 1.0,
) -> PathData:
    """Convert an SVG ``<rect>`` element to PathData."""
    try:
        x = float(elem.get("x", "0"))
        y = float(elem.get("y", "0"))
        w = float(elem.get("width", "0"))
        h = float(elem.get("height", "0"))
        rx = float(elem.get("rx", "0"))
        ry = float(elem.get("ry", "0"))
    except ValueError:
        rx = ry = 0.0
        x = y = w = h = 0.0

    if rx == 0 and ry == 0:
        # Simple rectangle via moveTo/lnTo/close
        corners = [(x, y), (x + w, y), (x + w, y + h), (x, y + h)]
        tc = [_apply_matrix(transform_matrix, cx, cy) for cx, cy in corners]
        commands = [f"moveTo {tc[0][0]} {tc[0][1]}"]
        for px, py in tc[1:]:
            commands.append(f"lnTo {px} {py}")
        commands.append("close")
    else:
        # Rounded rect via arc approximation at corners
        if ry == 0:
            ry = rx
        if rx == 0:
            rx = ry
        # Clamp radii
        rx = min(rx, w / 2)
        ry = min(ry, h / 2)

        # Build path with rounded corners using cubic beziers
        # Start at top-left after radius
        points: list[tuple[float, float]] = [
            (x + rx, y),           # top-left after radius
            (x + w - rx, y),       # top-right before radius
            (x + w, y + ry),       # top-right after radius
            (x + w, y + h - ry),   # bottom-right before radius
            (x + w - rx, y + h),   # bottom-right after radius
            (x + rx, y + h),       # bottom-left before radius
            (x, y + h - ry),       # bottom-left after radius
            (x, y + ry),           # top-left before radius
        ]

        # Transform all points
        tp = [_apply_matrix(transform_matrix, px, py) for px, py in points]

        # Approximate corner with cubic bezier (kappa constant)
        kappa = 0.5522847498  # Magic number for circular arc approximation

        commands = [f"moveTo {tp[0][0]} {tp[0][1]}"]

        # Top-right corner
        tr_ctrl1 = _apply_matrix(transform_matrix, x + w - rx * (1 - kappa), y)
        tr_ctrl2 = _apply_matrix(transform_matrix, x + w, y + ry * (1 - kappa))
        commands.append(f"cubicBezTo {tr_ctrl1[0]} {tr_ctrl1[1]} {tr_ctrl2[0]} {tr_ctrl2[1]} {tp[1][0]} {tp[1][1]}")
        commands.append(f"lnTo {tp[2][0]} {tp[2][1]}")

        # Bottom-right corner
        br_ctrl1 = _apply_matrix(transform_matrix, x + w, y + h - ry * (1 - kappa))
        br_ctrl2 = _apply_matrix(transform_matrix, x + w - rx * (1 - kappa), y + h)
        commands.append(f"cubicBezTo {br_ctrl1[0]} {br_ctrl1[1]} {br_ctrl2[0]} {br_ctrl2[1]} {tp[3][0]} {tp[3][1]}")
        commands.append(f"lnTo {tp[4][0]} {tp[4][1]}")

        # Bottom-left corner
        bl_ctrl1 = _apply_matrix(transform_matrix, x + rx * (1 - kappa), y + h)
        bl_ctrl2 = _apply_matrix(transform_matrix, x, y + h - ry * (1 - kappa))
        commands.append(f"cubicBezTo {bl_ctrl1[0]} {bl_ctrl1[1]} {bl_ctrl2[0]} {bl_ctrl2[1]} {tp[5][0]} {tp[5][1]}")
        commands.append(f"lnTo {tp[6][0]} {tp[6][1]}")

        # Top-left corner
        tl_ctrl1 = _apply_matrix(transform_matrix, x, y + ry * (1 - kappa))
        tl_ctrl2 = _apply_matrix(transform_matrix, x + rx * (1 - kappa), y)
        commands.append(f"cubicBezTo {tl_ctrl1[0]} {tl_ctrl1[1]} {tl_ctrl2[0]} {tl_ctrl2[1]} {tp[7][0]} {tp[7][1]}")

        commands.append("close")

    fill_str = _get_attr(elem, "fill")
    stroke_str = _get_attr(elem, "stroke")
    sw_str = _get_attr(elem, "stroke-width")
    opacity_str = _get_attr(elem, "opacity")

    raw_fill = fill_str if fill_str is not None else parent_fill
    raw_stroke = stroke_str if stroke_str is not None else parent_stroke
    raw_sw = sw_str if sw_str is not None else parent_stroke_width

    fill_hex = _parse_svg_color(raw_fill)
    stroke_hex = _parse_svg_color(raw_stroke)
    stroke_width = 1.0
    if raw_sw is not None:
        try:
            stroke_width = float(re.sub(r"[a-zA-Z]+$", "", str(raw_sw)))
        except ValueError:
            stroke_width = 1.0

    if fill_str is None and parent_fill is None and fill_hex is None:
        fill_hex = "000000"

    elem_opacity = _parse_svg_opacity(opacity_str)
    effective_opacity = parent_opacity * elem_opacity

    name = elem.get("id", f"SVG_Rect_{idx}")

    return PathData(
        d="\n".join(commands),
        fill=fill_hex,
        stroke=stroke_hex,
        stroke_width=stroke_width,
        name=name,
    )


def _convert_circle_element(
    elem: Any,
    transform_matrix: list[float],
    idx: int,
    parent_fill: str | None = None,
    parent_stroke: str | None = None,
    parent_stroke_width: float | None = None,
    parent_opacity: float = 1.0,
) -> PathData:
    """Convert an SVG ``<circle>`` element to PathData (4-arc approximation)."""
    try:
        cx = float(elem.get("cx", "0"))
        cy = float(elem.get("cy", "0"))
        r = float(elem.get("r", "0"))
    except ValueError:
        cx = cy = r = 0.0

    if r <= 0:
        return PathData(d="", name=elem.get("id", f"SVG_Circle_{idx}"))

    # Approximate circle with 4 cubic beziers (kappa constant)
    kappa = 0.5522847498

    # Four quadrants
    commands: list[str] = []
    start = (cx + r, cy)
    sx, sy = _apply_matrix(transform_matrix, start[0], start[1])
    commands.append(f"moveTo {sx} {sy}")

    # Arc from (cx+r, cy) to (cx, cy+r) — top-right quadrant
    p1 = _apply_matrix(transform_matrix, cx + r, cy - r * kappa)
    p2 = _apply_matrix(transform_matrix, cx + r * kappa, cy - r)
    p3 = _apply_matrix(transform_matrix, cx, cy - r)
    commands.append(f"cubicBezTo {p1[0]} {p1[1]} {p2[0]} {p2[1]} {p3[0]} {p3[1]}")

    # Arc from (cx, cy-r) to (cx-r, cy) — top-left quadrant
    p1 = _apply_matrix(transform_matrix, cx - r * kappa, cy - r)
    p2 = _apply_matrix(transform_matrix, cx - r, cy - r * kappa)
    p3 = _apply_matrix(transform_matrix, cx - r, cy)
    commands.append(f"cubicBezTo {p1[0]} {p1[1]} {p2[0]} {p2[1]} {p3[0]} {p3[1]}")

    # Arc from (cx-r, cy) to (cx, cy+r) — bottom-left quadrant
    p1 = _apply_matrix(transform_matrix, cx - r, cy + r * kappa)
    p2 = _apply_matrix(transform_matrix, cx - r * kappa, cy + r)
    p3 = _apply_matrix(transform_matrix, cx, cy + r)
    commands.append(f"cubicBezTo {p1[0]} {p1[1]} {p2[0]} {p2[1]} {p3[0]} {p3[1]}")

    # Arc from (cx, cy+r) to (cx+r, cy) — bottom-right quadrant
    p1 = _apply_matrix(transform_matrix, cx + r * kappa, cy + r)
    p2 = _apply_matrix(transform_matrix, cx + r, cy + r * kappa)
    p3 = _apply_matrix(transform_matrix, cx + r, cy)
    commands.append(f"cubicBezTo {p1[0]} {p1[1]} {p2[0]} {p2[1]} {p3[0]} {p3[1]}")

    commands.append("close")

    fill_str = _get_attr(elem, "fill")
    stroke_str = _get_attr(elem, "stroke")
    sw_str = _get_attr(elem, "stroke-width")

    raw_fill = fill_str if fill_str is not None else parent_fill
    raw_stroke = stroke_str if stroke_str is not None else parent_stroke
    raw_sw = sw_str if sw_str is not None else parent_stroke_width

    fill_hex = _parse_svg_color(raw_fill)
    stroke_hex = _parse_svg_color(raw_stroke)
    stroke_width = 1.0
    if raw_sw is not None:
        try:
            stroke_width = float(re.sub(r"[a-zA-Z]+$", "", str(raw_sw)))
        except ValueError:
            stroke_width = 1.0

    if fill_str is None and parent_fill is None and fill_hex is None:
        fill_hex = "000000"

    name = elem.get("id", f"SVG_Circle_{idx}")

    return PathData(
        d="\n".join(commands),
        fill=fill_hex,
        stroke=stroke_hex,
        stroke_width=stroke_width,
        name=name,
    )


def _convert_ellipse_element(
    elem: Any,
    transform_matrix: list[float],
    idx: int,
    parent_fill: str | None = None,
    parent_stroke: str | None = None,
    parent_stroke_width: float | None = None,
    parent_opacity: float = 1.0,
) -> PathData:
    """Convert an SVG ``<ellipse>`` element to PathData."""
    try:
        cx = float(elem.get("cx", "0"))
        cy = float(elem.get("cy", "0"))
        rx = float(elem.get("rx", "0"))
        ry = float(elem.get("ry", "0"))
    except ValueError:
        cx = cy = rx = ry = 0.0

    if rx <= 0 or ry <= 0:
        return PathData(d="", name=elem.get("id", f"SVG_Ellipse_{idx}"))

    kappa = 0.5522847498

    commands: list[str] = []
    start = (cx + rx, cy)
    sx, sy = _apply_matrix(transform_matrix, start[0], start[1])
    commands.append(f"moveTo {sx} {sy}")

    # Top-right quadrant
    p1 = _apply_matrix(transform_matrix, cx + rx, cy - ry * kappa)
    p2 = _apply_matrix(transform_matrix, cx + rx * kappa, cy - ry)
    p3 = _apply_matrix(transform_matrix, cx, cy - ry)
    commands.append(f"cubicBezTo {p1[0]} {p1[1]} {p2[0]} {p2[1]} {p3[0]} {p3[1]}")

    # Top-left quadrant
    p1 = _apply_matrix(transform_matrix, cx - rx * kappa, cy - ry)
    p2 = _apply_matrix(transform_matrix, cx - rx, cy - ry * kappa)
    p3 = _apply_matrix(transform_matrix, cx - rx, cy)
    commands.append(f"cubicBezTo {p1[0]} {p1[1]} {p2[0]} {p2[1]} {p3[0]} {p3[1]}")

    # Bottom-left quadrant
    p1 = _apply_matrix(transform_matrix, cx - rx, cy + ry * kappa)
    p2 = _apply_matrix(transform_matrix, cx - rx * kappa, cy + ry)
    p3 = _apply_matrix(transform_matrix, cx, cy + ry)
    commands.append(f"cubicBezTo {p1[0]} {p1[1]} {p2[0]} {p2[1]} {p3[0]} {p3[1]}")

    # Bottom-right quadrant
    p1 = _apply_matrix(transform_matrix, cx + rx * kappa, cy + ry)
    p2 = _apply_matrix(transform_matrix, cx + rx, cy + ry * kappa)
    p3 = _apply_matrix(transform_matrix, cx + rx, cy)
    commands.append(f"cubicBezTo {p1[0]} {p1[1]} {p2[0]} {p2[1]} {p3[0]} {p3[1]}")

    commands.append("close")

    fill_str = _get_attr(elem, "fill")
    stroke_str = _get_attr(elem, "stroke")
    sw_str = _get_attr(elem, "stroke-width")

    raw_fill = fill_str if fill_str is not None else parent_fill
    raw_stroke = stroke_str if stroke_str is not None else parent_stroke
    raw_sw = sw_str if sw_str is not None else parent_stroke_width

    fill_hex = _parse_svg_color(raw_fill)
    stroke_hex = _parse_svg_color(raw_stroke)
    stroke_width = 1.0
    if raw_sw is not None:
        try:
            stroke_width = float(re.sub(r"[a-zA-Z]+$", "", str(raw_sw)))
        except ValueError:
            stroke_width = 1.0

    if fill_str is None and parent_fill is None and fill_hex is None:
        fill_hex = "000000"

    name = elem.get("id", f"SVG_Ellipse_{idx}")

    return PathData(
        d="\n".join(commands),
        fill=fill_hex,
        stroke=stroke_hex,
        stroke_width=stroke_width,
        name=name,
    )


def _convert_line_element(
    elem: Any,
    transform_matrix: list[float],
    idx: int,
    parent_stroke: str | None = None,
    parent_stroke_width: float | None = None,
    parent_opacity: float = 1.0,
) -> PathData:
    """Convert an SVG ``<line>`` element to PathData."""
    try:
        x1 = float(elem.get("x1", "0"))
        y1 = float(elem.get("y1", "0"))
        x2 = float(elem.get("x2", "0"))
        y2 = float(elem.get("y2", "0"))
    except ValueError:
        x1 = y1 = x2 = y2 = 0.0

    sx, sy = _apply_matrix(transform_matrix, x1, y1)
    ex, ey = _apply_matrix(transform_matrix, x2, y2)

    commands = [f"moveTo {sx} {sy}", f"lnTo {ex} {ey}"]

    stroke_str = _get_attr(elem, "stroke")
    sw_str = _get_attr(elem, "stroke-width")

    raw_stroke = stroke_str if stroke_str is not None else parent_stroke
    raw_sw = sw_str if sw_str is not None else parent_stroke_width

    stroke_hex = _parse_svg_color(raw_stroke) or "000000"
    stroke_width = 1.0
    if raw_sw is not None:
        try:
            stroke_width = float(re.sub(r"[a-zA-Z]+$", "", str(raw_sw)))
        except ValueError:
            stroke_width = 1.0

    name = elem.get("id", f"SVG_Line_{idx}")

    return PathData(
        d="\n".join(commands),
        fill=None,
        stroke=stroke_hex,
        stroke_width=stroke_width,
        name=name,
    )


def _convert_poly_element(
    elem: Any,
    tag: str,
    transform_matrix: list[float],
    idx: int,
    parent_fill: str | None = None,
    parent_stroke: str | None = None,
    parent_stroke_width: float | None = None,
    parent_opacity: float = 1.0,
) -> PathData | None:
    """Convert an SVG ``<polygon>`` or ``<polyline>`` element to PathData."""
    points_str = elem.get("points")
    if not points_str:
        return None

    # Parse points: "x1,y1 x2,y2 ..." or "x1 y1 x2 y2 ..."
    points: list[tuple[float, float]] = []
    for m in re.finditer(r"([+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?)"
                         r"[,\s]+"
                         r"([+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?)",
                         points_str):
        try:
            px, py = float(m.group(1)), float(m.group(2))
            points.append((px, py))
        except ValueError:
            continue

    if len(points) < 2:
        return None

    tp = [_apply_matrix(transform_matrix, px, py) for px, py in points]

    commands = [f"moveTo {tp[0][0]} {tp[0][1]}"]
    for px, py in tp[1:]:
        commands.append(f"lnTo {px} {py}")

    is_polygon = tag.endswith("polygon") or tag == "polygon"
    if is_polygon:
        commands.append("close")

    fill_str = _get_attr(elem, "fill")
    stroke_str = _get_attr(elem, "stroke")
    sw_str = _get_attr(elem, "stroke-width")

    raw_fill = fill_str if fill_str is not None else parent_fill
    raw_stroke = stroke_str if stroke_str is not None else parent_stroke
    raw_sw = sw_str if sw_str is not None else parent_stroke_width

    fill_hex = _parse_svg_color(raw_fill)
    stroke_hex = _parse_svg_color(raw_stroke)
    stroke_width = 1.0
    if raw_sw is not None:
        try:
            stroke_width = float(re.sub(r"[a-zA-Z]+$", "", str(raw_sw)))
        except ValueError:
            stroke_width = 1.0

    # Polyline default: no fill; polygon default: fill black
    if not is_polygon:
        if fill_str is None and parent_fill is None:
            fill_hex = None
    else:
        if fill_str is None and parent_fill is None and fill_hex is None:
            fill_hex = "000000"

    prefix = "SVG_Polygon" if is_polygon else "SVG_Polyline"
    name = elem.get("id", f"{prefix}_{idx}")

    return PathData(
        d="\n".join(commands),
        fill=fill_hex,
        stroke=stroke_hex,
        stroke_width=stroke_width,
        name=name,
    )


# ---------------------------------------------------------------------------
# Recursive SVG tree walker
# ---------------------------------------------------------------------------

def _walk_svg_tree(
    elem: Any,
    transform_matrix: list[float],
    counter: list[int],
    parent_fill: str | None = None,
    parent_stroke: str | None = None,
    parent_stroke_width: float | None = None,
    parent_opacity: float = 1.0,
) -> list[PathData]:
    """Recursively walk the SVG element tree and convert each shape."""
    import xml.etree.ElementTree as ET

    results: list[PathData] = []

    # Strip namespace from tag
    tag = elem.tag
    if "}" in tag:
        tag = tag.split("}", 1)[1]
    tag_lower = tag.lower()

    # Compute cumulative transform
    local_transforms = _parse_transform(elem.get("transform"))
    if local_transforms:
        local_matrix = _concat_matrices(local_transforms)
        combined = _concat_matrices([transform_matrix, local_matrix])
    else:
        combined = transform_matrix

    # Inherit presentation attributes
    fill_attr = _get_attr(elem, "fill")
    stroke_attr = _get_attr(elem, "stroke")
    sw_attr = _get_attr(elem, "stroke-width")
    opacity_attr = _get_attr(elem, "opacity")

    current_fill = fill_attr if fill_attr is not None else parent_fill
    current_stroke = stroke_attr if stroke_attr is not None else parent_stroke
    current_sw = sw_attr if sw_attr is not None else parent_stroke_width
    current_opacity = parent_opacity * _parse_svg_opacity(opacity_attr)

    # Skip unsupported elements at this level
    if _element_has_unsupported_features(elem):
        return results

    if tag_lower == "path":
        pd = _convert_path_element(elem, combined, counter[0], current_fill, current_stroke, current_sw, current_opacity)
        if pd and pd.d:
            results.append(pd)
            counter[0] += 1
    elif tag_lower == "rect":
        pd = _convert_rect_element(elem, combined, counter[0], current_fill, current_stroke, current_sw, current_opacity)
        if pd.d:
            results.append(pd)
            counter[0] += 1
    elif tag_lower == "circle":
        pd = _convert_circle_element(elem, combined, counter[0], current_fill, current_stroke, current_sw, current_opacity)
        if pd.d:
            results.append(pd)
            counter[0] += 1
    elif tag_lower == "ellipse":
        pd = _convert_ellipse_element(elem, combined, counter[0], current_fill, current_stroke, current_sw, current_opacity)
        if pd.d:
            results.append(pd)
            counter[0] += 1
    elif tag_lower == "line":
        pd = _convert_line_element(elem, combined, counter[0], current_stroke, current_sw, current_opacity)
        if pd.d:
            results.append(pd)
            counter[0] += 1
    elif tag_lower in ("polygon", "polyline"):
        pd = _convert_poly_element(elem, tag_lower, combined, counter[0], current_fill, current_stroke, current_sw, current_opacity)
        if pd and pd.d:
            results.append(pd)
            counter[0] += 1

    # Recurse into children (e.g. <g> groups)
    for child in elem:
        results.extend(_walk_svg_tree(
            child, combined, counter,
            current_fill, current_stroke, current_sw, current_opacity,
        ))

    return results


# ---------------------------------------------------------------------------
# Public: pure conversion
# ---------------------------------------------------------------------------

def svg_to_drawingml(
    svg_path: str | None = None,
    svg_content: str | None = None,
) -> list[PathData]:
    """Convert SVG to DrawingML path data without creating a PPTX file.

    Parameters
    ----------
    svg_path : str, optional
        Path to an SVG file.
    svg_content : str, optional
        SVG content as a string.

    Returns
    -------
    list[PathData]
        One :class:`PathData` per SVG shape element.
    """
    import xml.etree.ElementTree as ET

    if svg_path:
        tree = ET.parse(svg_path)
        root = tree.getroot()
    elif svg_content:
        root = ET.fromstring(svg_content)
    else:
        raise ValueError("Either svg_path or svg_content must be provided")

    # Check for unsupported features
    if _svg_has_unsupported_features(root):
        raise ValueError(
            "SVG contains features that cannot be converted to DrawingML "
            "(filters, gradients, clip-paths, masks, text, images, symbols, "
            "or patterns). Use import_svg_as_image() as a fallback."
        )

    counter = [0]
    identity = [1, 0, 0, 1, 0, 0]
    return _walk_svg_tree(root, identity, counter)


# ---------------------------------------------------------------------------
# DrawingML XML builder helpers
# ---------------------------------------------------------------------------

def _build_cust_geom_xml(path_data: PathData, path_w: int, path_h: int) -> Any:
    """Build ``<a:custGeom>`` XML element from PathData."""
    from lxml import etree

    custGeom = etree.Element(f"{{{_NS_A}}}custGeom")

    # Empty required sub-elements
    etree.SubElement(custGeom, f"{{{_NS_A}}}avLst")
    etree.SubElement(custGeom, f"{{{_NS_A}}}gdLst")
    etree.SubElement(custGeom, f"{{{_NS_A}}}ahLst")
    etree.SubElement(custGeom, f"{{{_NS_A}}}cxnLst")

    rect = etree.SubElement(custGeom, f"{{{_NS_A}}}rect")
    rect.set("l", "0")
    rect.set("t", "0")
    rect.set("r", "0")
    rect.set("b", "0")

    pathLst = etree.SubElement(custGeom, f"{{{_NS_A}}}pathLst")
    path_elem = etree.SubElement(pathLst, f"{{{_NS_A}}}path")
    path_elem.set("w", str(path_w))
    path_elem.set("h", str(path_h))

    # Parse DrawingML command string
    for cmd_line in path_data.d.split("\n"):
        cmd_line = cmd_line.strip()
        if not cmd_line:
            continue
        parts = cmd_line.split()
        cmd_name = parts[0]

        if cmd_name == "moveTo" and len(parts) >= 3:
            moveTo = etree.SubElement(path_elem, f"{{{_NS_A}}}moveTo")
            pt = etree.SubElement(moveTo, f"{{{_NS_A}}}pt")
            pt.set("x", str(int(round(float(parts[1])))))
            pt.set("y", str(int(round(float(parts[2])))))

        elif cmd_name == "lnTo" and len(parts) >= 3:
            lnTo = etree.SubElement(path_elem, f"{{{_NS_A}}}lnTo")
            pt = etree.SubElement(lnTo, f"{{{_NS_A}}}pt")
            pt.set("x", str(int(round(float(parts[1])))))
            pt.set("y", str(int(round(float(parts[2])))))

        elif cmd_name == "cubicBezTo" and len(parts) >= 7:
            bez = etree.SubElement(path_elem, f"{{{_NS_A}}}cubicBezTo")
            for j in range(3):
                pt = etree.SubElement(bez, f"{{{_NS_A}}}pt")
                pt.set("x", str(int(round(float(parts[1 + j * 2])))))
                pt.set("y", str(int(round(float(parts[2 + j * 2])))))

        elif cmd_name == "close":
            etree.SubElement(path_elem, f"{{{_NS_A}}}close")

    return custGeom


def _build_fill_xml(fill_hex: str | None, opacity: float = 1.0) -> Any | None:
    """Build ``<a:solidFill>`` or ``<a:noFill>`` XML element."""
    from lxml import etree

    if fill_hex is None:
        noFill = etree.Element(f"{{{_NS_A}}}noFill")
        return noFill

    sf = etree.Element(f"{{{_NS_A}}}solidFill")
    clr = etree.SubElement(sf, f"{{{_NS_A}}}srgbClr")
    clr.set("val", fill_hex.upper())
    if opacity < 1.0:
        alpha = etree.SubElement(clr, f"{{{_NS_A}}}alpha")
        alpha.set("val", str(int(round(opacity * 100000))))
    return sf


def _build_stroke_xml(
    stroke_hex: str | None,
    stroke_width: float,
    dasharray: str | None = None,
    linecap: str | None = None,
    linejoin: str | None = None,
) -> Any | None:
    """Build ``<a:ln>`` XML element for stroke properties."""
    from lxml import etree

    if stroke_hex is None and stroke_width <= 0:
        return None

    ln = etree.Element(f"{{{_NS_A}}}ln")
    ln.set("w", str(_pt_to_emu(stroke_width)))

    if stroke_hex:
        sf = etree.SubElement(ln, f"{{{_NS_A}}}solidFill")
        clr = etree.SubElement(sf, f"{{{_NS_A}}}srgbClr")
        clr.set("val", stroke_hex.upper())
    else:
        etree.SubElement(ln, f"{{{_NS_A}}}noFill")

    # Dash style
    if dasharray:
        prst_dash = _svg_dasharray_to_preset(dasharray)
        if prst_dash:
            pd = etree.SubElement(ln, f"{{{_NS_A}}}prstDash")
            pd.set("val", prst_dash)

    # Line cap
    if linecap:
        cap_map = {"butt": "flat", "round": "rnd", "square": "sq"}
        ln.set("cap", cap_map.get(linecap, "flat"))

    # Line join
    if linejoin:
        join_map = {"round": "round", "bevel": "bevel", "miter": "miter"}
        if linejoin == "miter":
            mj = etree.SubElement(ln, f"{{{_NS_A}}}miter")
            mj.set("lim", "800000")
        elif linejoin in join_map:
            j = etree.SubElement(ln, f"{{{_NS_A}}}{join_map[linejoin]}")

    return ln


def _svg_dasharray_to_preset(dasharray: str) -> str | None:
    """Convert an SVG stroke-dasharray to the closest DrawingML preset dash."""
    # Parse dasharray values
    try:
        values = [float(x) for x in re.split(r"[\s,]+", dasharray.strip()) if x]
    except ValueError:
        return None

    if not values:
        return None

    # Normalize: consider just the pattern shape
    n = len(values)
    if n == 1:
        return "dash" if values[0] > 3 else "dot"
    elif n == 2:
        if values[0] > values[1] * 2:
            return "lgDash"
        elif values[0] < values[1]:
            return "dot"
        else:
            return "dash"
    elif n >= 3:
        # dash-dot pattern
        return "dashDot"

    return "dash"


# ---------------------------------------------------------------------------
# Presentation helpers
# ---------------------------------------------------------------------------

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


def _compute_shape_bounds(path_data_list: list[PathData]) -> tuple[float, float, float, float]:
    """Compute the bounding box of all PathData objects.

    Returns ``(min_x, min_y, max_x, max_y)`` in SVG coordinate space.
    """
    min_x = float("inf")
    min_y = float("inf")
    max_x = float("-inf")
    max_y = float("-inf")

    for pd in path_data_list:
        for cmd_line in pd.d.split("\n"):
            parts = cmd_line.strip().split()
            if not parts:
                continue
            cmd = parts[0]
            if cmd in ("moveTo", "lnTo") and len(parts) >= 3:
                try:
                    x, y = float(parts[1]), float(parts[2])
                    min_x = min(min_x, x)
                    min_y = min(min_y, y)
                    max_x = max(max_x, x)
                    max_y = max(max_y, y)
                except ValueError:
                    continue
            elif cmd == "cubicBezTo" and len(parts) >= 7:
                try:
                    for j in range(3):
                        x, y = float(parts[1 + j * 2]), float(parts[2 + j * 2])
                        min_x = min(min_x, x)
                        min_y = min(min_y, y)
                        max_x = max(max_x, x)
                        max_y = max(max_y, y)
                except ValueError:
                    continue

    if min_x == float("inf"):
        return 0.0, 0.0, 100.0, 100.0

    return min_x, min_y, max_x, max_y


def _add_shape_to_sp_tree(
    sp_tree: Any,
    path_data: PathData,
    offset_x: int,
    offset_y: int,
    cx: int,
    cy: int,
    path_w: int,
    path_h: int,
    fill_override: str | None = None,
    stroke_override: str | None = None,
    stroke_width_override: float | None = None,
) -> None:
    """Add a single ``<p:sp>`` with custom geometry to the shape tree."""
    from lxml import etree

    sp = etree.SubElement(sp_tree, f"{{{_NS_P}}}sp")
    nvSpPr = etree.SubElement(sp, f"{{{_NS_P}}}nvSpPr")
    cNvPr = etree.SubElement(nvSpPr, f"{{{_NS_P}}}cNvPr")
    cNvPr.set("id", "0")
    cNvPr.set("name", path_data.name)
    etree.SubElement(nvSpPr, f"{{{_NS_P}}}cNvSpPr")
    etree.SubElement(nvSpPr, f"{{{_NS_P}}}nvPr")

    spPr = etree.SubElement(sp, f"{{{_NS_P}}}spPr")

    # Transform (children of spPr use the 'a' namespace)
    xfrm = etree.SubElement(spPr, f"{{{_NS_A}}}xfrm")
    off = etree.SubElement(xfrm, f"{{{_NS_A}}}off")
    off.set("x", str(offset_x))
    off.set("y", str(offset_y))
    ext = etree.SubElement(xfrm, f"{{{_NS_A}}}ext")
    ext.set("cx", str(max(cx, 1)))
    ext.set("cy", str(max(cy, 1)))

    # Custom geometry
    custGeom = _build_cust_geom_xml(path_data, path_w, path_h)
    spPr.append(custGeom)

    # Fill
    fill_hex = fill_override if fill_override is not None else path_data.fill
    fill_elem = _build_fill_xml(fill_hex)
    if fill_elem is not None:
        spPr.append(fill_elem)

    # Stroke
    stroke_hex = stroke_override if stroke_override is not None else path_data.stroke
    sw = stroke_width_override if stroke_width_override is not None else path_data.stroke_width
    stroke_elem = _build_stroke_xml(stroke_hex, sw)
    if stroke_elem is not None:
        spPr.append(stroke_elem)

    # Text body (required by python-pptx for proper round-tripping)
    txBody = etree.SubElement(sp, f"{{{_NS_P}}}txBody")
    bodyPr = etree.SubElement(txBody, f"{{{_NS_A}}}bodyPr")
    bodyPr.set("rtlCol", "0")
    bodyPr.set("anchor", "ctr")
    etree.SubElement(txBody, f"{{{_NS_A}}}lstStyle")
    p = etree.SubElement(txBody, f"{{{_NS_A}}}p")
    pPr = etree.SubElement(p, f"{{{_NS_A}}}pPr")
    pPr.set("algn", "ctr")


def _add_group_to_sp_tree(
    sp_tree: Any,
    path_data_list: list[PathData],
    group_name: str,
    left_emu: int,
    top_emu: int,
    width_emu: int,
    height_emu: int,
    svg_vb_x: float,
    svg_vb_y: float,
    svg_vb_w: float,
    svg_vb_h: float,
    fill_override: str | None = None,
    stroke_override: str | None = None,
    stroke_width_override: float | None = None,
) -> None:
    """Add a ``<p:grpSp>`` containing all path shapes to the shape tree."""
    from lxml import etree

    # Compute overall bounds of all paths
    min_x, min_y, max_x, max_y = _compute_shape_bounds(path_data_list)

    svg_content_w = max_x - min_x if max_x > min_x else svg_vb_w
    svg_content_h = max_y - min_y if max_y > min_y else svg_vb_h

    # Scale factor: SVG units → EMU
    scale_x = width_emu / svg_vb_w if svg_vb_w > 0 else 1
    scale_y = height_emu / svg_vb_h if svg_vb_h > 0 else 1

    grpSp = etree.SubElement(sp_tree, f"{{{_NS_P}}}grpSp")

    # Group shape non-visual properties
    nvGrpSpPr = etree.SubElement(grpSp, f"{{{_NS_P}}}nvGrpSpPr")
    cNvPr = etree.SubElement(nvGrpSpPr, f"{{{_NS_P}}}cNvPr")
    cNvPr.set("id", "0")
    cNvPr.set("name", group_name)
    etree.SubElement(nvGrpSpPr, f"{{{_NS_P}}}cNvGrpSpPr")
    etree.SubElement(nvGrpSpPr, f"{{{_NS_P}}}nvPr")

    # Group shape properties
    grpSpPr = etree.SubElement(grpSp, f"{{{_NS_P}}}grpSpPr")

    xfrm = etree.SubElement(grpSpPr, f"{{{_NS_A}}}xfrm")
    off = etree.SubElement(xfrm, f"{{{_NS_A}}}off")
    off.set("x", str(left_emu))
    off.set("y", str(top_emu))
    ext = etree.SubElement(xfrm, f"{{{_NS_A}}}ext")
    ext.set("cx", str(max(width_emu, 1)))
    ext.set("cy", str(max(height_emu, 1)))

    chOff = etree.SubElement(xfrm, f"{{{_NS_A}}}chOff")
    chOff.set("x", str(_pt_to_emu(svg_vb_x)))
    chOff.set("y", str(_pt_to_emu(svg_vb_y)))
    chExt = etree.SubElement(xfrm, f"{{{_NS_A}}}chExt")
    chExt.set("cx", str(_pt_to_emu(svg_vb_w)))
    chExt.set("cy", str(_pt_to_emu(svg_vb_h)))

    # Add each path as a child shape
    for pd in path_data_list:
        if not pd.d.strip():
            continue

        # Convert SVG coordinates to child coordinate space (points)
        # The child coordinate system is defined by chOff/chExt
        # We need to convert our SVG-unit coordinates into this space
        # Since chOff/chExt maps the viewBox, and our coordinates are in
        # SVG units, we just need to convert to EMU using the viewBox scale

        path_w = _pt_to_emu(svg_vb_w)
        path_h = _pt_to_emu(svg_vb_h)

        # Each shape's offset is at the group origin (0,0 in child coords)
        # because the shape's path coordinates are already in SVG space
        child_offset_x = _pt_to_emu(svg_vb_x)
        child_offset_y = _pt_to_emu(svg_vb_y)
        child_cx = _pt_to_emu(svg_vb_w)
        child_cy = _pt_to_emu(svg_vb_h)

        # Scale path coordinates from SVG units to child EMU space
        scale_to_emu_x = _pt_to_emu(1)  # 1 SVG unit = 1 pt in child coords
        scale_to_emu_y = _pt_to_emu(1)

        # Transform path coordinates
        scaled_pd = _scale_path_data(pd, scale_to_emu_x, scale_to_emu_y)

        _add_shape_to_sp_tree(
            grpSp,
            scaled_pd,
            child_offset_x,
            child_offset_y,
            child_cx,
            child_cy,
            path_w,
            path_h,
            fill_override,
            stroke_override,
            stroke_width_override,
        )


def _scale_path_data(pd: PathData, sx: float, sy: float) -> PathData:
    """Scale all coordinates in PathData by (sx, sy)."""
    scaled_commands: list[str] = []
    for cmd_line in pd.d.split("\n"):
        parts = cmd_line.strip().split()
        if not parts:
            continue
        cmd = parts[0]
        if cmd in ("moveTo", "lnTo") and len(parts) >= 3:
            x = float(parts[1]) * sx
            y = float(parts[2]) * sy
            scaled_commands.append(f"{cmd} {x} {y}")
        elif cmd == "cubicBezTo" and len(parts) >= 7:
            coords = []
            for j in range(3):
                coords.append(float(parts[1 + j * 2]) * sx)
                coords.append(float(parts[2 + j * 2]) * sy)
            scaled_commands.append(
                f"cubicBezTo {coords[0]} {coords[1]} {coords[2]} {coords[3]} {coords[4]} {coords[5]}"
            )
        elif cmd == "close":
            scaled_commands.append("close")
        else:
            scaled_commands.append(cmd_line)

    return PathData(
        d="\n".join(scaled_commands),
        fill=pd.fill,
        stroke=pd.stroke,
        stroke_width=pd.stroke_width,
        name=pd.name,
    )


# ---------------------------------------------------------------------------
# Public: import_svg
# ---------------------------------------------------------------------------

def import_svg(
    prs_or_path: Any,
    slide_index: int,
    *,
    svg_path: str | None = None,
    svg_content: str | None = None,
    left: float = 1.0,
    top: float = 1.0,
    width: float | None = None,
    height: float | None = None,
    name: str | None = None,
    fill: str | None = None,
    stroke: str | None = None,
    stroke_width: float | None = None,
) -> str:
    """Import an SVG file or content string as native PowerPoint shapes.

    Converts SVG path data to ``<a:custGeom>`` custom geometry shapes that
    are fully editable in PowerPoint.  If the SVG contains features that
    cannot be represented in DrawingML (filters, gradients, text, etc.),
    a :class:`ValueError` is raised — use :func:`import_svg_as_image` as a
    fallback in that case.

    When the SVG contains multiple shape elements, they are grouped inside
    a ``<p:grpSp>`` element.

    Parameters
    ----------
    prs_or_path : Presentation or str
        A python-pptx Presentation object or a file path.
    slide_index : int
        Zero-based slide index.
    svg_path : str, optional
        Path to an SVG file.
    svg_content : str, optional
        SVG content as a string.
    left : float
        Left position in inches.  Default 1.0.
    top : float
        Top position in inches.  Default 1.0.
    width : float, optional
        Width in inches.  If not specified, uses SVG viewBox width (treating
        SVG units as points).
    height : float, optional
        Height in inches.  If not specified, uses SVG viewBox height.
    name : str, optional
        Base name for the shape(s).  Default ``"SVG_<index>"``.
    fill : str, optional
        Override fill colour (hex string).  Overrides SVG fill for all shapes.
    stroke : str, optional
        Override stroke colour (hex string).  Overrides SVG stroke for all shapes.
    stroke_width : float, optional
        Override stroke width in points.

    Returns
    -------
    str
        The group shape name (or single shape name if only one path).
    """
    import xml.etree.ElementTree as ET

    if svg_path is None and svg_content is None:
        raise ValueError("Either svg_path or svg_content must be provided")

    # Parse SVG
    if svg_path:
        tree = ET.parse(svg_path)
        root = tree.getroot()
    else:
        root = ET.fromstring(svg_content)  # type: ignore[arg-type]

    # Check for unsupported features
    if _svg_has_unsupported_features(root):
        raise ValueError(
            "SVG contains features that cannot be converted to DrawingML "
            "(filters, gradients, clip-paths, masks, text, images, symbols, "
            "or patterns). Use import_svg_as_image() as a fallback."
        )

    # Get viewBox
    vb_x, vb_y, vb_w, vb_h = _parse_svg_viewbox(root)

    # Determine target dimensions
    if width is None:
        width = vb_w / 72.0  # treat SVG units as points, convert to inches
    if height is None:
        height = vb_h / 72.0

    # Convert to EMU
    left_emu = _inch_to_emu(left)
    top_emu = _inch_to_emu(top)
    width_emu = _inch_to_emu(width)
    height_emu = _inch_to_emu(height)

    # Extract path data
    counter = [0]
    identity = [1, 0, 0, 1, 0, 0]
    path_data_list = _walk_svg_tree(root, identity, counter)

    if not path_data_list:
        raise ValueError("No convertible shape elements found in SVG")

    # Assign names with SVG_ prefix
    for i, pd in enumerate(path_data_list):
        if not pd.name.startswith("SVG_"):
            pd.name = f"SVG_{pd.name}"
        # Ensure uniqueness
        if name:
            pd.name = f"{name}_{i}"

    # Open presentation
    is_path = not _is_presentation(prs_or_path)
    path = prs_or_path if is_path else None
    prs = _open_prs(prs_or_path)
    try:
        slide = prs.slides[slide_index]
        sp_tree = slide.shapes._spTree

        group_name = name or f"SVG_Group_{len(slide.shapes)}"

        if len(path_data_list) == 1:
            # Single shape — no group needed
            pd = path_data_list[0]
            single_name = name or pd.name

            # Scale path coordinates from SVG units to EMU
            scale_x = width_emu / vb_w if vb_w > 0 else 1
            scale_y = height_emu / vb_h if vb_h > 0 else 1
            scaled_pd = _scale_path_data(pd, scale_x, scale_y)

            path_w = width_emu
            path_h = height_emu

            _add_shape_to_sp_tree(
                sp_tree,
                scaled_pd,
                left_emu,
                top_emu,
                width_emu,
                height_emu,
                path_w,
                path_h,
                fill,
                stroke,
                stroke_width,
            )

            # Fix the name
            for sp in sp_tree.findall(f"{{{_NS_P}}}sp"):
                cNvPr = sp.find(f".//{{{_NS_P}}}cNvPr")
                if cNvPr is not None and cNvPr.get("name") == pd.name:
                    cNvPr.set("name", single_name)
                    break

            return single_name
        else:
            # Multiple shapes — use group
            _add_group_to_sp_tree(
                sp_tree,
                path_data_list,
                group_name,
                left_emu,
                top_emu,
                width_emu,
                height_emu,
                vb_x,
                vb_y,
                vb_w,
                vb_h,
                fill,
                stroke,
                stroke_width,
            )

            return group_name

    finally:
        _save_prs(prs, path)


# ---------------------------------------------------------------------------
# Public: import_svg_as_image
# ---------------------------------------------------------------------------

def import_svg_as_image(
    prs_or_path: Any,
    slide_index: int,
    *,
    svg_path: str,
    left: float = 1.0,
    top: float = 1.0,
    width: float | None = None,
    height: float | None = None,
    dpi: int = 150,
    name: str | None = None,
) -> str:
    """Import an SVG as a rasterised PNG image (fallback path).

    Renders the SVG to a PNG using Cairo (via ``cairosvg``) or Pillow,
    then adds it as a picture shape.

    Parameters
    ----------
    prs_or_path : Presentation or str
        A python-pptx Presentation object or a file path.
    slide_index : int
        Zero-based slide index.
    svg_path : str
        Path to an SVG file.
    left : float
        Left position in inches.  Default 1.0.
    top : float
        Top position in inches.  Default 1.0.
    width : float, optional
        Width in inches.  If not specified, uses SVG viewBox.
    height : float, optional
        Height in inches.  If not specified, uses SVG viewBox.
    dpi : int
        Resolution for rasterisation.  Default 150.
    name : str, optional
        Shape name.  Default ``"SVG_Image_<index>"``.

    Returns
    -------
    str
        The picture shape name.
    """
    import os
    import tempfile
    import xml.etree.ElementTree as ET

    # Get SVG dimensions
    tree = ET.parse(svg_path)
    root = tree.getroot()
    vb_x, vb_y, vb_w, vb_h = _parse_svg_viewbox(root)

    if width is None:
        width = vb_w / 72.0
    if height is None:
        height = vb_h / 72.0

    # Render SVG to PNG
    png_data: bytes | None = None

    # Try cairosvg first (best quality)
    try:
        import cairosvg  # type: ignore[import-untyped]
        png_data = cairosvg.svg2png(
            url=svg_path,
            output_width=width * dpi,
            output_height=height * dpi,
        )
    except ImportError:
        pass
    except Exception:
        pass

    # Fallback: Pillow with svg support
    if png_data is None:
        try:
            from PIL import Image  # type: ignore[import-untyped]

            # Try Pillow's SVG support (requires pypotrace or similar)
            # Most likely won't work without additional deps
            png_path = os.path.join(
                tempfile.gettempdir(),
                f"svg_import_{os.getpid()}.png",
            )

            # Use Wand (ImageMagick) if available
            try:
                from wand.image import Image as WandImage  # type: ignore[import-untyped]
                with WandImage(filename=svg_path) as img:
                    img.resolution = dpi
                    img.resize(int(width * dpi), int(height * dpi))
                    img.save(filename=png_path)
                with open(png_path, "rb") as f:
                    png_data = f.read()
                os.unlink(png_path)
            except ImportError:
                pass

        except ImportError:
            pass

    if png_data is None:
        raise RuntimeError(
            "Cannot render SVG to PNG. Install cairosvg (pip install cairosvg) "
            "or Wand (pip install Wand) for SVG rasterisation support."
        )

    # Write PNG to temp file and add as picture
    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
        tmp.write(png_data)
        tmp_path = tmp.name

    try:
        prs = _open_prs(prs_or_path)
        try:
            slide = prs.slides[slide_index]
            _name = name or f"SVG_Image_{len(slide.shapes)}"

            left_emu = _inch_to_emu(left)
            top_emu = _inch_to_emu(top)
            width_emu = _inch_to_emu(width)
            height_emu = _inch_to_emu(height)

            pic = slide.shapes.add_picture(
                tmp_path, left_emu, top_emu, width_emu, height_emu,
            )
            pic.name = _name

            return _name
        finally:
            if isinstance(prs_or_path, str):
                prs.save(prs_or_path)
    finally:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass


# ---------------------------------------------------------------------------
# Public: list_svg_shapes
# ---------------------------------------------------------------------------

def list_svg_shapes(
    prs_or_path: Any,
    slide_index: int,
) -> list[dict]:
    """List all shapes on a slide that were imported from SVG.

    Identifies SVG-imported shapes by their name prefix ``"SVG_"``.

    Parameters
    ----------
    prs_or_path : Presentation or str
        A python-pptx Presentation object or a file path.
    slide_index : int
        Zero-based slide index.

    Returns
    -------
    list[dict]
        Each dict has keys: ``name``, ``shape_type``, ``left``, ``top``,
        ``width``, ``height``, ``has_custom_geometry``.
    """
    prs = _open_prs(prs_or_path)
    try:
        slide = prs.slides[slide_index]
        results: list[dict] = []

        for shape in slide.shapes:
            if not shape.name.startswith("SVG_"):
                continue

            # Check for custom geometry
            has_cust_geom = False
            try:
                elem = shape._element
                sp_pr = elem.find(f".//{{{_NS_A}}}custGeom")
                has_cust_geom = sp_pr is not None
            except Exception:
                pass

            results.append({
                "name": shape.name,
                "shape_type": "group" if shape.shape_type == 6 else "shape",
                "left": shape.left,
                "top": shape.top,
                "width": shape.width,
                "height": shape.height,
                "has_custom_geometry": has_cust_geom,
            })

        return results
    finally:
        pass  # Read-only operation
