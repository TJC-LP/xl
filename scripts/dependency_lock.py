#!/usr/bin/env python3
"""Record/check Mill's resolved Maven artifacts without changing dependency selection."""
import argparse
import difflib
import hashlib
import json
from pathlib import Path
import re
import subprocess
import sys
from urllib.parse import unquote
import xml.etree.ElementTree as ET


TARGETS = (
    "resolvedMvnDeps", "resolvedRunMvnDeps", "scalaCompilerClasspath",
    "scalacPluginClasspath", "scalaDocClasspath", "scalaDocPluginClasspath",
    "annotationProcessorsResolvedMvnDeps", "prepareOffline",
)
SELECTOR = "__.{" + ",".join(TARGETS) + "}"


def show(repo, selector, meta=False):
    command = [str(repo / "mill"), "-i"]
    if meta:
        command += ["--meta-level", "1"]
    command += ["show", selector]
    result = subprocess.run(command, cwd=repo, text=True, stdout=subprocess.PIPE, check=True)
    values = json.loads(result.stdout)
    # Mill unwraps a single selected task; wildcard selections with several tasks are objects.
    return values if isinstance(values, dict) else {selector: values}


def resolve(repo):
    result = {}
    for prefix, selector, meta in (
        ("project", SELECTOR, False),
        ("project", "__.generatorDeps", False),
        ("build", SELECTOR, True),
    ):
        values = show(repo, selector, meta)
        if not values:
            raise ValueError("No dependency targets resolved: " + prefix + "/" + selector)
        result.update({prefix + "/" + key: value for key, value in values.items()})
    return result


def artifact(ref):
    """Decode a Mill PathRef, then derive a Maven identity from its repository layout.

    Use the filename too: classifiers/extensions must not alias the ordinary JAR.
    Accept any Coursier HTTP(S) repository, not only Central or one user's cache root.
    Fail closed for unmanaged files or unrecognised layouts instead of omitting them.
    """
    match = re.fullmatch(r"(?:qref|ref):v\d+:[^:]+:(.+)", ref)
    if not match:
        raise ValueError("Unsupported Mill PathRef: " + str(ref))
    path = Path(match[1])
    parts = path.as_posix().split("/")
    try:
        index = next(i for i, part in enumerate(parts) if part in ("https", "http"))
    except StopIteration:
        raise ValueError("Artifact is not in a Coursier HTTP(S) cache: " + str(path))
    # Repository paths can have prefixes (maven2, artifactory/repo, ...). Find
    # group by its adjacent POM's XML, rather than treating that prefix as a group.
    if len(parts) < index + 6:
        raise ValueError("Unsupported Maven artifact path: " + str(path))
    name, version, filename = (unquote(part) for part in parts[-3:])
    pom = path.parent / (name + "-" + version + ".pom")
    # Coursier retains POMs alongside the JARs. Resolve group from the POM while
    # checking it against the cache path; never trust a filename alone as identity.
    root = ET.parse(pom).getroot()
    ns = {"m": "http://maven.apache.org/POM/4.0.0"}
    def pom_text(key):
        return root.findtext("/".join("m:" + p for p in key.split("/")), namespaces=ns) or root.findtext(key)
    group = pom_text("groupId") or pom_text("parent/groupId")
    if not group or "${" in group:
        raise ValueError("Cannot identify Maven group in " + str(pom))
    suffix = "/" + group.replace(".", "/") + "/" + name + "/" + version + "/" + filename
    if not unquote(path.as_posix()).endswith(suffix):
        raise ValueError("Maven coordinates do not match cache path: " + str(path))
    if not filename.startswith(name + "-" + version + ".") and not filename.startswith(name + "-" + version + "-"):
        raise ValueError("Unsupported Maven artifact filename: " + filename)
    return ":".join((group, name, version)) + "/" + filename, path


def snapshot(repo, resolved):
    artifacts, targets, hashed = {}, {}, {}
    for target, refs in sorted(resolved.items()):
        if not isinstance(refs, list):
            raise ValueError("Expected dependency list for " + target)
        entries = []
        for ref in refs:
            # NativeImageModule.prepareOffline also returns XL class directories
            # and resource directories. These are first-party build inputs, not
            # Maven artifacts. Never omit an unmanaged JAR or an external path.
            match = re.fullmatch(r"(?:qref|ref):v\d+:[^:]+:(.+)", ref)
            if target.endswith(".prepareOffline") and match:
                path = Path(match[1])
                try:
                    local = path.relative_to(repo).as_posix()
                except ValueError:
                    local = ""
                if re.fullmatch(r"(?:out/[^/]+/compile\.dest/classes|[^/]+/(?:compile-)?resources)", local):
                    if not path.is_file():
                        continue
            if ref not in hashed:
                identity, path = artifact(ref)
                digest = hashlib.sha256()
                with path.open("rb") as stream:
                    for block in iter(lambda: stream.read(1024 * 1024), b""):
                        digest.update(block)
                hashed[ref] = identity, digest.hexdigest()
            identity, checksum = hashed[ref]
            if identity in artifacts and artifacts[identity] != checksum:
                raise ValueError("Conflicting bytes for " + identity)
            artifacts[identity] = checksum
            entries.append(identity)
        # Preserve classpath order as well as membership; paths/hash-cache tokens are omitted.
        targets[target] = entries
    if not artifacts:
        raise ValueError("Resolved dependency inventory is empty")
    return {
        "formatVersion": 1,
        "millVersion": (repo / ".mill-version").read_text().strip(),
        "targets": targets,
        "artifacts": dict(sorted(artifacts.items())),
    }


def encode(value):
    return json.dumps(value, indent=2, ensure_ascii=True) + "\n"


def check(lock, actual):
    expected = lock.read_text()
    current = encode(actual)
    if expected != current:
        sys.stderr.writelines(difflib.unified_diff(
            expected.splitlines(keepends=True), current.splitlines(keepends=True),
            fromfile=str(lock), tofile="resolved dependencies",
        ))
        raise ValueError("Dependency lock drift. Investigate before running the update command.")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("command", choices=("check", "update"))
    args = parser.parse_args()
    repo = Path(__file__).resolve().parents[1]
    lock = repo / "deps.lock.json"
    try:
        actual = snapshot(repo, resolve(repo))
        if args.command == "update":
            lock.write_text(encode(actual))
        else:
            check(lock, actual)
        print("Dependency lock " + args.command + ": " + str(len(actual["artifacts"]))
              + " artifacts, " + str(len(actual["targets"])) + " targets")
    except (OSError, ValueError, ET.ParseError, subprocess.CalledProcessError) as error:
        print("Dependency lock failed: " + str(error), file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
