package com.tjclp.xl.xml

/**
 * Token kinds produced by [[XmlPullParser.next]].
 *
 * The token itself is a bare discriminator; the parser's cursor accessors ([[XmlPullParser.name]],
 * [[XmlPullParser.attrs]], [[XmlPullParser.text]], ...) carry the payload and are valid only until
 * the next call to `next()` (SAX-style aliasing contract).
 *
 *   - `StartElement` — an element start tag (or the start half of `<e/>`; the parser synthesizes
 *     the matching `EndElement` on the following `next()`).
 *   - `EndElement` — an element end tag (real or synthesized for an empty-element tag).
 *   - `Text` — one maximal coalesced run of character data (entities and character references
 *     decoded; comments and processing instructions are consumed silently and do not split runs).
 *   - `CData` — one CDATA section (its own token, never merged with adjacent text).
 *   - `EndDocument` — end of input after the root element closed; sticky (every later `next()`
 *     returns it again).
 */
enum XmlToken derives CanEqual:
  case StartElement, EndElement, Text, CData, EndDocument
