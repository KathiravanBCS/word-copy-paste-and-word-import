# tables/merged-cells

Merged cells in both directions: a `colspan=3` banner row, a `rowspan=2` cell,
and a `colspan=2` cell in the row it reaches into.

The point of the fixture is the grid resolution. Row 3 has two `<td>`s but
occupies three grid columns, because the rowspan from row 2 takes the first
one. Without resolving that, column widths land on the wrong columns and the
grid column count comes out as 2.
