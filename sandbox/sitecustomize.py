"""
Loaded by every Python in the sandbox (site-packages/sitecustomize.py).

The one reader of this sandbox's output is a language model, and the one
thing it does with a failed subprocess is `print(e)`. CalledProcessError's
message stops at "returned non-zero exit status 1", so a node script that
threw a TypeError, or soffice refusing a file, reached the model as an exit
code and nothing else - and its next attempt was a guess. Recorded twice in
one afternoon before this existed.

So the captured streams are appended to the message when the caller kept
them (capture_output=True / stderr=PIPE). Nothing changes for a caller that
did not capture; nothing changes for code that inspects e.stderr itself.
"""

import subprocess

_original_str = subprocess.CalledProcessError.__str__


def _explain(self):
    message = _original_str(self)

    for label, stream in (("stderr", self.stderr), ("stdout", self.stdout)):
        if not stream:
            continue
        text = stream.decode("utf-8", "replace") if isinstance(stream, bytes) else str(stream)
        text = text.strip()
        if text:
            message += "\n[%s]\n%s" % (label, text[-4000:])

    return message


subprocess.CalledProcessError.__str__ = _explain


# ---------------------------------------------------------------------------
# Documents left in /tmp are delivered.
#
# The other way a deck went missing: the model built it, previewed it, wrote
# "download it here" - and never printed the file's data URI, because writing
# to /tmp felt like enough. /tmp is discarded with the container, so the link
# pointed at nothing. Recorded on 2026-09-14 (gemma-4-26b-it).
#
# So at exit every document the run left in /tmp is printed as a named data
# URI unless the program already printed one with that name - the same thing
# OpenAI's sandbox does by listing /mnt/data. Only document types: a PNG is
# a preview the program prints when it wants it seen, and a PDF next to a
# .pptx or .docx of the same name is LibreOffice's conversion step, not a
# second deliverable.
# ---------------------------------------------------------------------------

import atexit
import base64
import os
import sys

_DELIVERABLE_MIMES = {
    ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ".pdf": "application/pdf",
    ".csv": "text/csv",
}

_MAX_AUTO_DELIVERED_BYTES = 20 * 1024 * 1024

_printed = []


class _RecordingStdout:
    """Remembers what the program printed, so a file it delivered itself is not delivered twice."""

    def __init__(self, wrapped):
        self._wrapped = wrapped

    def write(self, text):
        if "name=" in text:
            _printed.append(text)
        return self._wrapped.write(text)

    def __getattr__(self, name):
        return getattr(self._wrapped, name)


sys.stdout = _RecordingStdout(sys.stdout)


def _already_printed(basename):
    from urllib.parse import quote

    needles = ("name=" + basename + ";", "name=" + quote(basename) + ";", "name=" + quote(basename, safe="") + ";")
    return any(needle in chunk for chunk in _printed for needle in needles)


def _deliver_documents_left_in_tmp():
    try:
        entries = sorted(os.scandir("/tmp"), key=lambda e: e.name)
    except OSError:
        return

    files = {e.name: e for e in entries if e.is_file(follow_symlinks=False)}
    stems_with_office = {
        os.path.splitext(n)[0] for n in files if os.path.splitext(n)[1].lower() in (".pptx", ".docx")
    }

    for name, entry in files.items():
        stem, ext = os.path.splitext(name)
        mime = _DELIVERABLE_MIMES.get(ext.lower())
        if mime is None or name.startswith("."):
            continue
        if ext.lower() == ".pdf" and stem in stems_with_office:
            continue  # the conversion step, not the deliverable
        if _already_printed(name):
            continue
        try:
            if entry.stat().st_size > _MAX_AUTO_DELIVERED_BYTES:
                continue
            with open(entry.path, "rb") as handle:
                data = handle.read()
        except OSError:
            continue
        if not data:
            continue
        # The name is a data URI parameter: percent-encoded, so a space or an
        # umlaut in "Anthropomorphisierung von LLMs.pptx" cannot end the
        # parameter early. The receiver decodes it back into the file name.
        from urllib.parse import quote
        sys.stdout.write("\ndata:%s;name=%s;base64,%s\n" % (mime, quote(name, safe=""), base64.b64encode(data).decode("ascii")))

    try:
        sys.stdout.flush()
    except Exception:
        pass


atexit.register(_deliver_documents_left_in_tmp)
