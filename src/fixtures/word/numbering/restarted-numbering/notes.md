# numbering/restarted-numbering

The list-identity test from the spec: `1. A / 2. B / 3. C` is one list, and the
`1. D` that follows the interrupting paragraph is a *new* list, not a
continuation.

The distinction cannot be made from the visible numbers alone — both start at
1. It comes from the `mso-list` declaration: the second run uses list `l1` with
override `lfo2`, so it is a different list even though it looks identical.
