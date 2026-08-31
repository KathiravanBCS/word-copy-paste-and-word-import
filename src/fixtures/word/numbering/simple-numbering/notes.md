# numbering/simple-numbering

Word's default three-level numbering ladder: decimal, lower-alpha,
lower-roman.

Two things to notice in the payload. First, level 1 declares *no*
`mso-level-number-format` at all — the absence means arabic, not "unknown", and
a parser that treats it as unknown loses the format. Second, item 3 returns to
level 1 after two deeper levels and must continue at 3, while the deeper levels
restart if they are re-entered.
