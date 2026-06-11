# Native-image metadata for xl-cli (Apache Batik)

GraalVM `native-image` auto-discovers these configs from the classpath
(`META-INF/native-image/<group>/<artifact>/`). They register what Batik 1.19
needs at runtime so the bundled rasterizer can work in native binaries on
platforms where GraalVM supports AWT (Linux; see GH-86):

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

Status: best-effort, derived from the batik 1.19 jars' actual service files and
resources (not from a tracing-agent run). AWT is unsupported in native images on
macOS/Windows, where the CLI falls back to the subprocess backends — verify the
Linux native binary on the next release build.
