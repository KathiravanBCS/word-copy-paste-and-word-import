# Fixture library

Each directory here is one Word clipboard payload and its expected output:

```
<category>/<name>/
  input.html            the payload, byte for byte
  expected-model.json   the projected canonical model
  expected.html         the rendered output plus the generated stylesheet
  notes.md              what this fixture tests, and why it is hard
```

Read `notes.md` first — it explains what Word actually sends and what the
correct handling is, which is usually not obvious from the markup.

## Provenance

These fixtures are **hand-authored** to reproduce the markup shapes Microsoft
Word emits: the `ProgId`/`Generator` meta pair, the Office namespaces, the
comment-wrapped `<style>` block with its `@font-face` / `@list` / `Mso*` rules,
the CF_HTML `StartFragment` markers, the `WordSection1` wrapper, the
`<![if !supportLists]>` marker blocks, the `<o:p>` paragraph marks and the
`mso-*` declarations.

They are not captures from a running copy of Word. Captures are better evidence
and are welcome — see `docs/TESTING.md` for how to add one from a real paste
using the clipboard lab.

## Blessing

```bash
UPDATE_FIXTURES=1 npm test
git diff src/fixtures/
```

Read the diff. A change here is either a regression or an intentional behaviour
change, and there is no third possibility.
