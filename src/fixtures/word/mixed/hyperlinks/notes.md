# mixed/hyperlinks

Four link targets Word produces, and one bookmark.

- An external `https:` link containing a **bold** run — the run boundary inside
  the link must survive, which is why hyperlinks live out of line in the model
  and are referenced by id from runs.
- A `mailto:` link.
- An internal `#_Toc…` jump, matched by the `<a name>` bookmark further down.
  `_Toc` bookmarks are marked `internal` so a consumer can drop Word's own
  plumbing without losing user-authored bookmarks.
- A `file:///` link, which is not an injection risk but cannot be followed from
  a web page. It is reported and rendered as an inert link carrying its original
  target, not as a link that silently does nothing.
