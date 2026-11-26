"""
evidence_archive.py
--------------------
Shared chain-of-custody helper for this project's scraping scripts.

Every HTML page (or other file) fetched from a source website is:
  1. Saved verbatim -- the exact raw bytes received over the wire, with
     no re-encoding or normalization -- to original_htmls/<city>/...
  2. Hashed with SHA-256, computed over those exact raw bytes.
  3. Logged to a per-city original_htmls/<city>/manifest.csv row, with
     the source URL, retrieval timestamp (UTC), local file path, and
     hash, so the acquisition context and integrity proof travel
     together.

This directly supports the Berkeley Protocol on Digital Open Source
Investigations' chain-of-custody guidance: keep the original material,
hash it at the moment of capture, and log the acquisition context (URL +
timestamp + method), so a disputed record can later be verified against
what was actually on the page.

Import this from any scraper in this project:
    from evidence_archive import archive_response, archive_bytes

Typical usage inside a fetch function, right after a 200 response:
    sha256_hex, local_path = archive_response(
        resp, city="donetsk", identifier="novosti_1234"
    )
"""

from __future__ import annotations

import csv
import hashlib
import os
import re
from datetime import datetime, timezone

ARCHIVE_ROOT = "original_htmls"

_SAFE_ID_RE = re.compile(r"[^A-Za-z0-9_.-]")

_MANIFEST_FIELDS = [
    "identifier", "source_url", "fetched_at_utc",
    "local_path", "sha256", "content_type", "byte_length",
]


def _safe_filename(identifier: str, ext: str) -> str:
    safe = _SAFE_ID_RE.sub("_", identifier).strip("_") or "unnamed"
    return f"{safe}{ext}"


def _manifest_path(city: str) -> str:
    return os.path.join(ARCHIVE_ROOT, city, "manifest.csv")


def _append_manifest(city: str, row: dict) -> None:
    path = _manifest_path(city)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    file_exists = os.path.exists(path)
    with open(path, "a", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=_MANIFEST_FIELDS)
        if not file_exists:
            writer.writeheader()
        writer.writerow(row)


def archive_bytes(raw_bytes: bytes, city: str, identifier: str, source_url: str,
                   ext: str = ".html", content_type: str = "text/html") -> tuple[str, str]:
    """
    Save raw_bytes verbatim under original_htmls/<city>/, compute its
    SHA-256, append a manifest row, and return (sha256_hex, local_path).

    IMPORTANT: raw_bytes must be the exact bytes received over the wire
    (e.g. requests' response.content), NOT response.text. Hashing
    already-decoded/re-encoded text hashes a copy, not the original --
    that defeats the purpose of a chain-of-custody hash, which is meant
    to prove the archived file matches what the server actually sent.
    """
    city_dir = os.path.join(ARCHIVE_ROOT, city)
    os.makedirs(city_dir, exist_ok=True)

    filename = _safe_filename(identifier, ext)
    local_path = os.path.join(city_dir, filename)

    # Don't silently overwrite an earlier capture of the same identifier
    # (e.g. a re-run, or a page whose content changed between runs) --
    # each capture is potential evidence in its own right, so keep both.
    if os.path.exists(local_path):
        base, ext_only = os.path.splitext(local_path)
        n = 2
        while os.path.exists(f"{base}__v{n}{ext_only}"):
            n += 1
        local_path = f"{base}__v{n}{ext_only}"

    with open(local_path, "wb") as f:
        f.write(raw_bytes)

    sha256_hex = hashlib.sha256(raw_bytes).hexdigest()

    _append_manifest(city, {
        "identifier": identifier,
        "source_url": source_url,
        "fetched_at_utc": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "local_path": local_path,
        "sha256": sha256_hex,
        "content_type": content_type,
        "byte_length": len(raw_bytes),
    })

    return sha256_hex, local_path


def archive_response(resp, city: str, identifier: str, ext: str = ".html",
                      content_type: str = "text/html") -> tuple[str, str]:
    """Convenience wrapper: archive a requests.Response's raw .content
    (the exact bytes received) under the response's own .url."""
    return archive_bytes(resp.content, city, identifier, resp.url,
                          ext=ext, content_type=content_type)


def hash_and_log_file(path: str, city: str, identifier: str, source_url: str,
                       content_type: str = "application/octet-stream") -> str:
    """
    For files that are already saved to disk in their own location (e.g. a
    downloaded .xlsx/.docx/.pdf attachment kept in temp/individual_lists/,
    rather than under original_htmls/) -- compute its SHA-256 without
    duplicating the file, write a "<path>.sha256" sidecar next to it, log a
    manifest row pointing at its existing location, and return the hex
    digest. Use this instead of archive_bytes() when you don't want a
    second copy of the file living under original_htmls/.
    """
    with open(path, "rb") as f:
        raw_bytes = f.read()
    sha256_hex = hashlib.sha256(raw_bytes).hexdigest()

    sidecar_path = path + ".sha256"
    with open(sidecar_path, "w", encoding="utf-8") as f:
        f.write(f"{sha256_hex}  {os.path.basename(path)}\n")

    _append_manifest(city, {
        "identifier": identifier,
        "source_url": source_url,
        "fetched_at_utc": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "local_path": path,
        "sha256": sha256_hex,
        "content_type": content_type,
        "byte_length": len(raw_bytes),
    })
    return sha256_hex


def city_from_url(url: str) -> str:
    """
    Derive a filesystem-safe 'city'/site identifier from a URL's host, for
    scripts that crawl many different domains (e.g. dozens of individual
    court sites under sudrf.ru) rather than working against one fixed city
    constant. Each distinct host gets its own original_htmls/<host>/
    folder and manifest, same as the per-city scripts use a literal city
    name -- keeping every source's chain-of-custody records separate and
    easy to audit independently.
    """
    from urllib.parse import urlparse
    netloc = urlparse(url).netloc
    safe = re.sub(r"[^A-Za-z0-9]+", "_", netloc).strip("_")
    return safe or "unknown_host"
