#!/usr/bin/env python3
from __future__ import annotations

import argparse
import subprocess
import sys
import time
from pathlib import Path

import uno


def _prop(name: str, value) -> object:
    p = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
    p.Name = name
    p.Value = value
    return p


def _connect_desktop(host: str, port: int, *, timeout_s: float) -> object:
    local_ctx = uno.getComponentContext()
    resolver = local_ctx.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local_ctx)
    deadline = time.time() + timeout_s
    last_err: Exception | None = None
    while time.time() < deadline:
        try:
            ctx = resolver.resolve(f"uno:socket,host={host},port={port};urp;StarOffice.ComponentContext")
            smgr = ctx.ServiceManager
            desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
            return desktop
        except Exception as e:  # pragma: no cover
            last_err = e
            time.sleep(0.2)
    raise RuntimeError(f"Failed to connect to LibreOffice UNO (last error: {last_err})")


def _file_url(path: Path) -> str:
    return path.resolve().as_uri()


def main() -> int:
    parser = argparse.ArgumentParser(description="Update TOC/fields in a DOCX using LibreOffice headless (UNO).")
    parser.add_argument("input", type=Path, help="Input .docx")
    parser.add_argument("--output", type=Path, default=None, help="Output .docx (default: overwrite input)")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=2002)
    parser.add_argument("--timeout-s", type=float, default=20.0)
    args = parser.parse_args()

    in_path: Path = args.input
    out_path: Path = args.output or in_path
    if not in_path.exists():
        print(f"Missing input: {in_path}", file=sys.stderr)
        return 2

    profile_dir = Path("/tmp") / f"lo-profile-{int(time.time())}"
    profile_dir.mkdir(parents=True, exist_ok=True)
    profile_url = profile_dir.resolve().as_uri()

    cmd = [
        "soffice",
        "--headless",
        "--nologo",
        "--nolockcheck",
        "--nodefault",
        "--norestore",
        "--invisible",
        f'--accept=socket,host={args.host},port={args.port};urp;',
        f"-env:UserInstallation={profile_url}",
    ]

    proc = subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    try:
        desktop = _connect_desktop(args.host, args.port, timeout_s=args.timeout_s)

        load_props = (
            _prop("Hidden", True),
            _prop("ReadOnly", False),
            # Make the importer explicit; in some environments LO may not auto-detect .docx reliably via UNO.
            _prop("FilterName", "MS Word 2007 XML"),
        )
        doc = None
        for _ in range(15):
            try:
                doc = desktop.loadComponentFromURL(_file_url(in_path), "_blank", 0, load_props)
            except Exception:
                doc = None
            if doc is not None:
                break
            time.sleep(0.2)
        if doc is None:
            print("Failed to load document via UNO (doc is None).", file=sys.stderr)
            return 3

        # Update indexes (TOC is an index) and fields.
        try:
            indexes = doc.getDocumentIndexes()
            for i in range(indexes.getCount()):
                indexes.getByIndex(i).update()
        except Exception:
            pass
        try:
            doc.refresh()
        except Exception:
            pass
        try:
            doc.getTextFields().refresh()
        except Exception:
            pass

        store_props = (
            _prop("FilterName", "MS Word 2007 XML"),
            _prop("Overwrite", True),
        )
        doc.storeToURL(_file_url(out_path), store_props)
        doc.close(True)
        try:
            desktop.terminate()
        except Exception:
            pass
    finally:
        proc.terminate()
        try:
            proc.wait(timeout=5)
        except Exception:
            proc.kill()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
