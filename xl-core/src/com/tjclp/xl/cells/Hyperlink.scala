package com.tjclp.xl.cells

/**
 * A cell's hyperlink target: an external URL (`https://…`, `mailto:…`, `file:…`) or an internal
 * location (`Sheet1!A1`, a defined name).
 *
 * A distinct type rather than a `String`, so that a positional `Some("text")` in `Cell`'s fourth
 * parameter does not compile: that slot held the removed `Cell.comment`, and a comment written
 * there would otherwise silently become a link (#891). Zero overhead: the representation is the
 * target string itself.
 *
 * `Hyperlink(target)` holds the target exactly as given. `Cell.withHyperlink(url: String)` is the
 * authoring chokepoint that strips the leading `#` of the `#Sheet1!A1` dialect (GH-406), and the
 * OOXML writer strips it again on emit, so a verbatim `Hyperlink("#Sheet1!A1")` still writes a
 * clean `location`.
 *
 * Members stay non-`inline` and are not export-forwarded: see the opaque-type note in
 * `addressing/ARef.scala`. `target` resolves through this companion's implicit scope.
 */
opaque type Hyperlink = String

object Hyperlink:
  /** A hyperlink to `target`, stored verbatim. */
  def apply(target: String): Hyperlink = target

  extension (h: Hyperlink)
    /** The link target as written: a URL or an internal location. */
    def target: String = h
