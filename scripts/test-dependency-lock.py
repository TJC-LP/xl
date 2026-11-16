#!/usr/bin/env python3
"""Adversarial checks for the dependency integrity gate; no network or JVM needed."""
import contextlib
import copy
import importlib.util
import io
import json
from pathlib import Path
import subprocess
import tempfile
import unittest
from unittest.mock import patch


SPEC = importlib.util.spec_from_file_location("dependency_lock", Path(__file__).with_name("dependency_lock.py"))
LOCK = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(LOCK)


class DependencyLockTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory(prefix="dependency lock ")
        self.addCleanup(self.tmp.cleanup)
        self.repo = Path(self.tmp.name) / "repo"
        self.repo.mkdir()
        (self.repo / ".mill-version").write_text("1.1.5\n")
        self.jar = self.maven("org.example", "lib_3", "1.0", b"tested library")
        self.refs = {"project/app.resolvedMvnDeps": [self.ref(self.jar)]}
        self.lock = self.repo / "deps.lock.json"
        self.expected = LOCK.snapshot(self.repo, self.refs)
        self.lock.write_text(LOCK.encode(self.expected))

    def maven(self, group, name, version, content, repository="maven2", classifier=""):
        directory = Path(self.tmp.name) / "cache/v1/https/repo.example" / repository / group.replace(".", "/") / name / version
        directory.mkdir(parents=True, exist_ok=True)
        jar = directory / (name + "-" + version + classifier + ".jar")
        jar.write_bytes(content)
        (directory / (name + "-" + version + ".pom")).write_text(
            '<project xmlns="http://maven.apache.org/POM/4.0.0"><parent>'
            '<groupId>' + group + '</groupId></parent><artifactId>' + name + '</artifactId></project>'
        )
        return jar

    def ref(self, path):
        return "qref:v1:cached-hash:" + str(path)

    def assert_drift(self, actual):
        with contextlib.redirect_stderr(io.StringIO()), self.assertRaisesRegex(ValueError, "drift"):
            LOCK.check(self.lock, actual)

    def test_unchanged_tree_and_bytes_pass(self):
        LOCK.check(self.lock, LOCK.snapshot(self.repo, self.refs))

    def test_changed_bytes_fail_even_with_same_mill_cache_hash(self):
        self.jar.write_bytes(b"replaced in cache")
        self.assert_drift(LOCK.snapshot(self.repo, self.refs))

    def test_changed_version_fails(self):
        other = self.maven("org.example", "lib_3", "1.1", b"new library")
        self.refs["project/app.resolvedMvnDeps"] = [self.ref(other)]
        self.assert_drift(LOCK.snapshot(self.repo, self.refs))

    def test_artifact_absent_from_lock_fails(self):
        other = self.maven("org.example", "transitive", "2.0", b"transitive")
        self.refs["project/app.resolvedMvnDeps"].append(self.ref(other))
        self.assert_drift(LOCK.snapshot(self.repo, self.refs))

    def test_removed_artifact_fails(self):
        self.refs["project/app.resolvedMvnDeps"] = []
        with self.assertRaisesRegex(ValueError, "empty"):
            LOCK.snapshot(self.repo, self.refs)

    def test_configuration_changes_cannot_hide_in_union(self):
        self.refs["project/app.resolvedRunMvnDeps"] = [self.ref(self.jar)]
        self.lock.write_text(LOCK.encode(LOCK.snapshot(self.repo, self.refs)))
        self.refs["project/app.resolvedMvnDeps"] = []
        self.assert_drift(LOCK.snapshot(self.repo, self.refs))

    def test_new_or_removed_target_fails(self):
        actual = copy.deepcopy(self.expected)
        actual["targets"]["project/new.test.resolvedMvnDeps"] = []
        self.assert_drift(actual)
        del actual["targets"]["project/app.resolvedMvnDeps"]
        del actual["targets"]["project/new.test.resolvedMvnDeps"]
        self.assert_drift(actual)

    def test_classpath_order_is_checked(self):
        other = self.maven("org.example", "other", "1.0", b"other library")
        self.refs["project/app.resolvedMvnDeps"].append(self.ref(other))
        self.lock.write_text(LOCK.encode(LOCK.snapshot(self.repo, self.refs)))
        self.refs["project/app.resolvedMvnDeps"].reverse()
        self.assert_drift(LOCK.snapshot(self.repo, self.refs))

    def test_mirror_cache_location_does_not_change_lock(self):
        mirror = self.maven("org.example", "lib_3", "1.0", self.jar.read_bytes(), repository="artifactory/central")
        actual = LOCK.snapshot(self.repo, {"project/app.resolvedMvnDeps": [self.ref(mirror)]})
        self.assertEqual(actual, self.expected)

    def test_conflicting_repository_bytes_fail(self):
        mirror = self.maven("org.example", "lib_3", "1.0", b"different", repository="other")
        self.refs["project/other.resolvedMvnDeps"] = [self.ref(mirror)]
        with self.assertRaisesRegex(ValueError, "Conflicting bytes"):
            LOCK.snapshot(self.repo, self.refs)

    def test_classifier_is_distinct(self):
        other = self.maven("org.example", "lib_3", "1.0", b"tests", classifier="-tests")
        self.refs["project/app.resolvedMvnDeps"].append(self.ref(other))
        self.assertEqual(len(LOCK.snapshot(self.repo, self.refs)["artifacts"]), 2)

    def test_missing_jar_or_pom_fails(self):
        self.jar.unlink()
        with self.assertRaises(OSError):
            LOCK.snapshot(self.repo, self.refs)
        self.jar.write_bytes(b"restored")
        self.jar.with_suffix(".pom").unlink()
        with self.assertRaises(OSError):
            LOCK.snapshot(self.repo, self.refs)

    def test_unknown_reference_or_unmanaged_jar_fails(self):
        for ref in ["unexpected", self.ref(self.repo / "lib.jar")]:
            with self.subTest(ref=ref), self.assertRaises(ValueError):
                LOCK.snapshot(self.repo, {"project/app.resolvedMvnDeps": [ref]})

    def test_only_known_first_party_offline_directories_are_omitted(self):
        self.refs["project/app.prepareOffline"] = [self.ref(self.repo / "out/app/compile.dest/classes"), self.ref(self.repo / "app/resources")]
        actual = LOCK.snapshot(self.repo, self.refs)
        self.assertEqual(actual["artifacts"], self.expected["artifacts"])
        (self.repo / "app").mkdir()
        (self.repo / "app/resources").write_bytes(b"not a directory")
        with self.assertRaises(ValueError):
            LOCK.snapshot(self.repo, self.refs)

    def test_resolution_failure_is_not_empty_success(self):
        failure = subprocess.CalledProcessError(1, ["mill"])
        with patch.object(LOCK.subprocess, "run", side_effect=failure), self.assertRaises(subprocess.CalledProcessError):
            LOCK.resolve(self.repo)
        self.assertEqual(json.loads(self.lock.read_text()), self.expected)


if __name__ == "__main__":
    unittest.main()
