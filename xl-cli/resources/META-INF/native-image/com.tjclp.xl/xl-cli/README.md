# Native-image metadata for xl-cli (Apache Batik + JDK Xerces)

GraalVM `native-image` auto-discovers these configs from the classpath
(`META-INF/native-image/<group>/<artifact>/`). They register what Batik 1.19
needs at runtime so the bundled rasterizer can work in native binaries on
platforms where GraalVM supports AWT (Linux; see GH-86), plus the JDK-internal
Xerces/Xalan message bundles (GH-349, see below):

- `reflect-config.json`
  - `org.apache.batik.css.parser.Parser`: loaded via `Class.forName` from the
    `org.w3c.css.sac.driver` key in
    `org/apache/batik/util/resources/XMLResourceDescriptor.properties`.
  - JDK Xerces JAXP factories: Batik 1.19 leaves `org.xml.sax.driver` unset and
    falls back to JAXP (`SAXParserFactory.newInstance()`).
  - `RegistryEntry`/`ImageWriter` implementations: instantiated reflectively by
    Batik's own service loader (`org.apache.batik.util.Service`) from the
    `META-INF/services` files enumerated below.
- `resource-config.json`
  - All Batik `.properties` (message bundles + `XMLResourceDescriptor.properties`,
    `dtdids.properties`), declared both as resources and as `ResourceBundle`s.
  - The two Batik service files actually shipped in batik-codec 1.19
    (`RegistryEntry`, `ImageWriter`). The `InterpreterFactory` service (Rhino,
    for scripted SVGs) is deliberately omitted — xl never emits `<script>`.

## JDK Xerces/Xalan message bundles (GH-349)

`resource-config.json` also registers nine JDK-internal message bundles: the
eight `com.sun.org.apache.xerces.internal.impl.msg.*Messages` bundles used by
SAX/DOM parsing and validation, plus the class-based
`com.sun.org.apache.xml.internal.res.XMLErrorResources`. `reflect-config.json`
registers the JDK Xerces JAXP factories, but without the bundles every Xerces
diagnostic on a native binary degraded to
`Could not load any resource bundle by com.sun...XMLMessages`, masking the
real parse error (GH-349, hit in the field on a style-bloated workbook).

- Base names verified against `jimage list` of the `java.xml` module (JDK 25).
- Registered by unqualified name (the legacy pre-module format) rather than the
  module-qualified new format (`{"module": "java.xml", "bundle": ...}`) —
  matching how the Batik bundles above are declared. If a native build ever
  logs `Could not find resource bundle` warnings for these names, switch to
  module-qualified registration.
- Guarded by the "Smoke test native binary (GH-349)" step in
  `.github/workflows/release.yml`, which runs the built binary against
  `xl-ooxml/test/resources/fixtures/malformed-workbook.xlsx` and asserts the
  real Xerces diagnostic (not the bundle-lookup failure) is reported. JVM-side
  message text is pinned by `WorkbookMetadataReaderSpec` (GH-349 test).

Status: best-effort, derived from the batik 1.19 jars' actual service files and
resources plus the JDK 25 `java.xml` jimage listing (not from a tracing-agent
run). AWT is unsupported in native images on macOS/Windows, where the CLI falls
back to the subprocess backends — the release smoke covers metadata reads on
every platform; verify rasterization on the Linux native binary on release
builds.
