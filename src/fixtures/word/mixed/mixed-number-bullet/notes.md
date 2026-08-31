# mixed/mixed-number-bullet

The mixed-list golden test:

```
I. Parent
    1. Child
        • Grandchild
            a. Great-grandchild
```

One list definition, four levels, alternating between numbered and bulleted.
Every level must keep its own marker type, format and glyph — a renderer that
picks `<ol>` or `<ul>` once for the whole hierarchy gets this wrong.
