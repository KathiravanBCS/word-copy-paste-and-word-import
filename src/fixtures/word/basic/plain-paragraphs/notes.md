# basic/plain-paragraphs

Baseline: text, a deliberately blank paragraph (`<o:p>&nbsp;</o:p>`), an
alignment, and indentation.

Asserts that Word's empty spacer paragraph survives as a paragraph rather than
being dropped, and that `margin-left`/`text-indent` reach the model in
canonical units.
