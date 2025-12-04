# Releasing XL to Maven Central

This document describes how to release new versions of XL to Maven Central.

## Prerequisites

Before your first release:
1. Sonatype Central account with verified `com.tjclp` namespace
2. GPG key uploaded to keyservers
3. GitHub Secrets configured (see below)

## Required GitHub Secrets

Configure these in `Settings > Secrets and variables > Actions`:

| Secret | Description |
|--------|-------------|
| `SONATYPE_USERNAME` | Sonatype Central username or user token |
| `SONATYPE_PASSWORD` | Sonatype Central password or user token |
| `PGP_SECRET` | Base64-encoded GPG private key |
| `PGP_PASSPHRASE` | GPG key passphrase |

Mill auto-imports these as `MILL_*` prefixed environment variables.

## Published Modules

| Module | Artifact ID | Description |
|--------|-------------|-------------|
| xl-core | `xl-core_3` | Core domain model, macros, DSL |
| xl-ooxml | `xl-ooxml_3` | OOXML readers and writers |
| xl-cats-effect | `xl-cats-effect_3` | IO interpreters, streaming |
| xl-evaluator | `xl-evaluator_3` | Formula parser and evaluator |

**Not published**: `xl-cli` (CLI binary), `xl-benchmarks` (JMH), `xl-testkit` (internal)

## Release Steps

### 1. Prepare the Release

```bash
# Ensure you're on main and up-to-date
git checkout main
git pull origin main

# Verify all tests pass
./mill __.test

# Verify formatting
./mill __.checkFormat
```

### 2. Create and Push Tag

```bash
# Create annotated tag
git tag -a v0.1.0 -m "Release 0.1.0"

# Push tag to trigger release workflow
git push origin v0.1.0
```

### 3. Monitor Release

1. Go to https://github.com/TJC-LP/xl/actions
2. Watch the "Release" workflow
3. Once complete, verify at https://central.sonatype.com

### 4. Verify Publication

After ~10-30 minutes, artifacts should be available:
- https://repo1.maven.org/maven2/com/tjclp/

## Version Numbering

We follow [Semantic Versioning](https://semver.org/):

| Type | Example | When |
|------|---------|------|
| MAJOR | 1.0.0 | Breaking API changes |
| MINOR | 0.2.0 | New features, backward compatible |
| PATCH | 0.1.1 | Bug fixes, backward compatible |
| Pre-release | 0.1.0-RC1 | Release candidates |

## One-Time Setup

### Sonatype Central Account

1. Register at https://central.sonatype.com/
2. Verify namespace ownership (`com.tjclp`)
3. Generate user token for CI

### GPG Key Setup

```bash
# Generate key (RSA 4096-bit)
gpg --full-generate-key

# Get key ID
gpg --list-secret-keys --keyid-format LONG
# Look for: sec rsa4096/ABCD1234EFGH5678

# Upload to keyservers
gpg --keyserver keyserver.ubuntu.com --send-keys ABCD1234EFGH5678
gpg --keyserver keys.openpgp.org --send-keys ABCD1234EFGH5678

# Export for GitHub Secret (base64 encoded)
gpg --export-secret-key -a ABCD1234EFGH5678 | base64 | pbcopy
```

## Usage in Downstream Projects

### Mill

```scala
def ivyDeps = Seq(
  mvn"com.tjclp::xl-core:0.1.0",
  mvn"com.tjclp::xl-cats-effect:0.1.0"
)
```

### SBT

```scala
libraryDependencies ++= Seq(
  "com.tjclp" %% "xl-core" % "0.1.0",
  "com.tjclp" %% "xl-cats-effect" % "0.1.0"
)
```

### Maven

```xml
<dependency>
  <groupId>com.tjclp</groupId>
  <artifactId>xl-core_3</artifactId>
  <version>0.1.0</version>
</dependency>
```

## Troubleshooting

### Release workflow fails with GPG error

- Verify `PGP_SECRET_BASE64` is correctly base64-encoded
- Verify `PGP_PASSPHRASE` matches the key passphrase
- Ensure key is uploaded to keyservers

### Artifacts not appearing on Maven Central

- Check Sonatype Central portal for deployment status
- Releases can take 10-30 minutes to sync
- Verify namespace is verified and active

### Version already exists

- Cannot overwrite released versions
- Bump version and create new tag

## Local Testing

```bash
# Build artifacts locally
./mill __.publishLocal

# Verify in local Ivy cache
ls ~/.ivy2/local/com/tjclp/

# Check generated POM
cat ~/.ivy2/local/com/tjclp/xl-core_3/0.1.0-SNAPSHOT/poms/xl-core_3.pom
```
