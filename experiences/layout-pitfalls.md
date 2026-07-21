# Layout Pitfalls

## TL;DR

Never hardcode absolute coordinates in layout constraints. Use relative constraints with canvas references and token values. Avoid contradictory constraints on the same variable. Mind the direction of gap constraints and the semantics of height vs. bottom in kiwisolver.

---

## Rule 1: Never use hardcoded absolute coordinates

Hardcoding pixel/point values breaks when the canvas size changes (e.g., 4:3 vs 16:9). Always reference canvas dimensions and token spacing values.

**Bad:**
```python
# Breaks on any canvas other than 16:9
constraint = (zone.left == 48)
constraint = (zone.width == 864)
```

**Good:**
```python
# Relative to canvas and token profile
constraint = (zone.left == canvas.left + tokens.margin_side)
constraint = (zone.width == canvas.width - 2 * tokens.margin_side)
```

This is the "BMC canvas bug" — constraints that work on one canvas silently produce wrong geometry on another.

---

## Rule 2: Avoid both `==` and `>=` on the same gap

Kiwi treats these as contradictory. If you set `gap == 24` and `gap >= 24`, the solver raises `UnsatisfiableConstraint` because the equality already implies the inequality, but the solver sees two constraints on the same variable with different strengths/types.

**Bad:**
```python
constraints = [
    (gap == tokens.gutter),   # equality
    (gap >= tokens.gutter),   # redundant, causes UnsatisfiableConstraint
]
```

**Good:**
```python
constraints = [
    (gap == tokens.gutter),   # equality alone is sufficient
]
```

If you need a minimum with flexibility, use a weak equality or a medium-strength `>=` without the `==`:

```python
constraints = [
    (gap >= tokens.gutter, kiwi.strength.medium),
]
```

---

## Rule 3: Gap constraint direction matters

The constraint `b.left - a.right == gap` is correct. The reverse `a.right - b.left == gap` produces a negative gap (b overlaps a from the left), which is almost never what you want.

**Bad:**
```python
# This means: a.right - b.left == gap → b is to the LEFT of a
(a.right - b.left == tokens.gutter)
```

**Good:**
```python
# This means: b.left - a.right == gap → b is to the RIGHT of a
(b.left - a.right == tokens.gutter)
```

Mnemonic: "next element's left minus previous element's right equals the gap."

---

## Rule 4: Cover hero zone needs both left AND width constraints

A common mistake is constraining only the height of a cover hero zone, leaving its horizontal position and width undefined. The solver will either fail or produce a zero-width zone.

**Bad:**
```python
constraints = [
    (hero.top == canvas.top),
    (hero.bottom == canvas.bottom - tokens.margin_bottom),
    # Missing: hero.left and hero.width → zero-width or unsolved
]
```

**Good:**
```python
constraints = [
    (hero.top == canvas.top),
    (hero.bottom == canvas.bottom - tokens.margin_bottom),
    (hero.left == canvas.left + tokens.margin_side),
    (hero.width == canvas.width - 2 * tokens.margin_side),
]
```

---

## Rule 5: Use bottom constraints instead of height with coefficients

Kiwi treats `height` as an independent variable. When you write `2 * child.height == parent.height`, kiwi may not propagate correctly because it doesn't automatically derive `height = bottom - top`. Instead, express the constraint in terms of `top` and `bottom`:

**Bad:**
```python
# Kiwi treats height as independent; coefficient math may not propagate
(2 * child.height == parent.height)
```

**Good:**
```python
# Express in terms of top/bottom which kiwi can chain through
(2 * child.bottom == parent.top + parent.bottom)
# This is equivalent to: child's vertical center == parent's vertical center
# and child's height == parent's height / 2
```

This pattern is especially important for vertically centering elements or distributing space proportionally.

---

## Rule 6: Always define all four edges of a zone

A zone with only three edges constrained will have one edge floating. The solver may assign it an arbitrary value (often 0), causing zones to collapse or overlap.

**Checklist for every zone:**
- `zone.left` is constrained
- `zone.right` is constrained (or `zone.width` + `zone.left`)
- `zone.top` is constrained
- `zone.bottom` is constrained (or `zone.height` + `zone.top`)

If you use `width`/`height`, ensure the opposite edge is also set:
```python
(zone.left == some_value),
(zone.width == some_other_value),
# zone.right is derived, but you must have both left and width
```
