"""Refresh the template's real PivotTables with LibreOffice Calc's DataPilot engine.
The resulting XLSX keeps the interactive PivotTable/filter structure for Excel.
"""
from __future__ import annotations

import os
import shutil
import socket
import subprocess
import sys
import tempfile
import time
from pathlib import Path


def _prop(name, value):
    from com.sun.star.beans import PropertyValue
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


def _uno_worker(path: str, port: int) -> None:
    import uno

    local_ctx = uno.getComponentContext()
    resolver = local_ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local_ctx
    )
    ctx = None
    last_error = None
    for _ in range(40):
        try:
            ctx = resolver.resolve(
                f"uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext"
            )
            break
        except Exception as exc:
            last_error = exc
            time.sleep(0.25)
    if ctx is None:
        raise RuntimeError(f"Could not connect to LibreOffice: {last_error}")

    desktop = ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.frame.Desktop", ctx
    )
    url = uno.systemPathToFileUrl(str(Path(path).resolve()))
    doc = desktop.loadComponentFromURL(
        url,
        "_blank",
        0,
        (_prop("Hidden", True), _prop("UpdateDocMode", 3)),
    )
    if doc is None:
        raise RuntimeError("LibreOffice could not open the workbook")

    try:
        refreshed = 0
        for sheet in doc.Sheets:
            tables = sheet.getDataPilotTables()
            for name in tables.getElementNames():
                tables.getByName(name).refresh()
                refreshed += 1

        # Keep Pivot value counts as numbers, never date-formatted values.
        # Pivot result blocks are small and start at A1 in the template.
        for sheet_name in ("R1 V", "R1S V", "R2 V", "R6 V", "R1 OCV", "R6 OCV"):
            if doc.Sheets.hasByName(sheet_name):
                sheet = doc.Sheets.getByName(sheet_name)
                # Column B is the Count - Task No value column.
                sheet.getCellRangeByName("B1:B1048576").NumberFormat = 0

        doc.calculateAll()
        doc.store()
        print(f"REFRESHED_PIVOTS={refreshed}")
    finally:
        doc.close(True)


def _free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        return int(s.getsockname()[1])


def refresh_workbook(path: str | Path) -> bool:
    """Refresh all PivotTables in-place. Returns False if LibreOffice is unavailable."""
    path = str(Path(path).resolve())
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    system_python = "/usr/bin/python3" if Path("/usr/bin/python3").exists() else None
    if not soffice or not system_python:
        return False

    port = _free_port()
    profile = Path(tempfile.mkdtemp(prefix="lo_profile_"))
    profile_url = profile.as_uri()
    proc = subprocess.Popen(
        [
            soffice,
            "--headless",
            f"-env:UserInstallation={profile_url}",
            f"--accept=socket,host=127.0.0.1,port={port};urp;StarOffice.ServiceManager",
            "--norestore",
            "--nodefault",
            "--nofirststartwizard",
        ],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    try:
        result = subprocess.run(
            [system_python, __file__, "--uno-worker", path, str(port)],
            capture_output=True,
            text=True,
            timeout=180,
        )
        if result.returncode != 0:
            raise RuntimeError(result.stderr.strip() or result.stdout.strip() or "Pivot refresh failed")
        return True
    finally:
        proc.terminate()
        try:
            proc.wait(timeout=10)
        except subprocess.TimeoutExpired:
            proc.kill()
        shutil.rmtree(profile, ignore_errors=True)


if __name__ == "__main__":
    if len(sys.argv) >= 2 and sys.argv[1] == "--uno-worker":
        _uno_worker(sys.argv[2], int(sys.argv[3]))
    elif len(sys.argv) >= 2:
        print("Refreshing", sys.argv[1])
        ok = refresh_workbook(sys.argv[1])
        print("OK" if ok else "LibreOffice is not available")
    else:
        print("Usage: pivot_refresh.py <xlsx>")
