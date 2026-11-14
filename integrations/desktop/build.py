"""Build an allowlisted MCPB using the exact verified public Panes SDK."""

import hashlib
import json
from pathlib import Path
import urllib.request
import zipfile

ROOT = Path(__file__).resolve().parent
SDK_NAME = "panes-0.1.3-py3-none-any.whl"
SDK_URL = "https://usepanes.com/sdk/" + SDK_NAME
SDK_SHA256 = "cc31e6c7d8803ad679312955fd1ae835bd0763c999629d4c76438d1304885e1e"
SOURCES = ("manifest.json", "pyproject.toml", "uv.lock", "src/server.py", "README.md")


def prepare_sdk():
    with urllib.request.urlopen(SDK_URL, timeout=30) as response:
        data = response.read(2 * 1024 * 1024)
    if hashlib.sha256(data).hexdigest() != SDK_SHA256:
        raise ValueError("Public SDK hash changed. Inspect it before updating the pin.")
    target = ROOT / "wheels" / SDK_NAME
    target.parent.mkdir(exist_ok=True)
    target.write_bytes(data)
    return target


def build():
    sdk = prepare_sdk()
    manifest = json.loads((ROOT / "manifest.json").read_text())
    version = manifest["version"]
    output = ROOT / "dist" / f"panes-{version}.mcpb"
    output.parent.mkdir(exist_ok=True)
    content = {name: (ROOT / name).read_bytes() for name in SOURCES}
    content["wheels/" + SDK_NAME] = sdk.read_bytes()
    content["LICENSE"] = (ROOT.parents[1] / "LICENSE").read_bytes()
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as bundle:
        for name, data in sorted(content.items()):
            info = zipfile.ZipInfo(name, date_time=(2026, 11, 14, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o100644 << 16
            bundle.writestr(info, data)
    digest = hashlib.sha256(output.read_bytes()).hexdigest()
    output.with_suffix(".mcpb.sha256").write_text(f"{digest}  {output.name}\n")
    print(json.dumps({"file": str(output), "sha256": digest, "sdk_sha256": SDK_SHA256}))


if __name__ == "__main__":
    import sys
    if sys.argv[1:] == ["--prepare-sdk"]:
        prepare_sdk()
    elif not sys.argv[1:]:
        build()
    else:
        raise SystemExit("Usage: python build.py [--prepare-sdk]")
