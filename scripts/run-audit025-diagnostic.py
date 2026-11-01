#!/usr/bin/env python3
"""Retain one bounded allocation-counter diagnostic; never rerun to get green."""
from pathlib import Path
import hashlib
import json
import os
import platform
import subprocess
import uuid


def main():
    root = Path(__file__).resolve().parents[1]
    assembly = root / "tests/Ofdrw.Net.Packaging.Tests/bin/Release/net10.0/Ofdrw.Net.Packaging.Tests.dll"
    core = assembly.with_name("Ofdrw.Net.Core.dll")
    font = root / "e2e/Ofdrw.Net.FontSubset.E2E/testdata/fonts/NotoSans-Regular.ttf"
    folder = Path(os.environ.get("OFDRW_AUDIT025_DIAGNOSTICS_DIRECTORY", root / "artifacts/audit025-diagnostics")).resolve()
    folder.mkdir(parents=True, exist_ok=True)
    output = folder / f"pressure-{platform.system().lower()}-{uuid.uuid4().hex}.json"
    env = os.environ.copy()
    env.setdefault("DOTNET_CLI_HOME", str(root / "artifacts/dotnet-home"))
    env.update(DOTNET_SKIP_FIRST_TIME_EXPERIENCE="1", DOTNET_CLI_TELEMETRY_OPTOUT="1")
    command = [env.get("DOTNET_HOST_PATH", "dotnet"), str(assembly), "--audit025-pressure", str(font), str(output)]
    run = subprocess.run(command, env=env, timeout=30, capture_output=True, text=True)
    # Keep the first raw result even when execution or an assertion fails.
    print(f"Audit025 pressure receipt: {output}", flush=True)
    if run.returncode:
        raise RuntimeError(f"Diagnostic exited {run.returncode}: {run.stdout}\n{run.stderr}")
    data = json.loads(output.read_text())
    digest = lambda path: hashlib.sha256(path.read_bytes()).hexdigest()
    if not data["Passed"] or data["Mode"] != "--audit025-pressure" or data["PressureSamplesRecorded"] != 128:
        raise RuntimeError("Incomplete diagnostic; raw receipt retained.")
    if data["Core"]["Sha256"] != digest(core) or data["FixtureSha256"] != digest(font):
        raise RuntimeError("Actual loaded Core/fixture identity mismatch.")
    source = os.environ.get("GITHUB_SHA")
    if source and not data["Core"]["InformationalVersion"].endswith("+" + source):
        raise RuntimeError("Actual Core source identity mismatch.")
    print(f"Actual Core SHA256: {digest(core)}; raw receipt SHA256: {digest(output)}", flush=True)
    print("Diagnostic completed; observed deltas do not attribute a previous Windows failure.", flush=True)


if __name__ == "__main__":
    main()
