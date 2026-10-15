#!/usr/bin/env python3
# ============================================================
# FETCH AFRO LINE LIST FROM SHAREPOINT
#
# Downloads the WHO AFRO polio line list --
#   AllPolioviruses_20250106.xlsx
# -- from the WHO AF-pep/GISWORKSPACE SharePoint site, via a sharing link,
# and saves it locally as:
#   data/afro_line_list.xlsx
#
# PORTABLE BY DESIGN: every path this script uses (config/, data/) is
# resolved relative to THIS FILE's own location, not to whatever folder
# you happen to run it from. Copy the whole folder this script lives in
# (fetch_afro_line_list.py + config/ + data/) to any machine, fill in
# config/secrets.env there, and it works the same way -- no hardcoded
# usernames or drive letters anywhere in this file.
#
# Uses the same app-only (client-credentials) Microsoft Graph flow as the
# other WHO AFRO tooling (im_workflow's upload_to_sharepoint.py,
# prepare_the_AFRO_SIA_Dashboard_input.py's SharePoint downloads): resolves
# the sharing link straight to a driveItem via Graph's /shares/{id}
# endpoint, then downloads its content. No browser, no MFA.
#
# ------------------------------------------------------------
# SETUP (one-time, on each new machine/station):
#   1) pip install requests
#   2) Copy config/secrets.env.example to config/secrets.env and fill in:
#        SHAREPOINT_TENANT_ID
#        SHAREPOINT_CLIENT_ID
#        SHAREPOINT_CLIENT_SECRET
#      These are the SAME app-registration credentials already used by
#      im_workflow / prepare_the_AFRO_SIA_Dashboard_input.py for this same
#      SharePoint site -- copy them from wherever you already keep those.
#      config/secrets.env is for YOUR EYES ONLY: never commit it, never
#      share it, never paste its contents into a chat. If this script's
#      folder is run for the first time and secrets.env is missing, it
#      will create it FROM the .example template automatically and stop,
#      so you have a file ready to fill in.
# ------------------------------------------------------------
#
# Usage (run from anywhere -- paths are resolved relative to this file):
#   python fetch_afro_line_list.py
#   python fetch_afro_line_list.py --output-name afro_line_list_2025.xlsx
#   python fetch_afro_line_list.py --share-url "https://...a-different-file..."
# ============================================================

import argparse
import os
import shutil
import sys
from datetime import datetime
from pathlib import Path

# Windows consoles frequently can't print accented characters unless the
# console codepage happens to be UTF-8 -- reconfigure so a stray print()
# falls back to '?'-style replacement instead of crashing the whole run.
try:
    sys.stdout.reconfigure(errors="replace")
    sys.stderr.reconfigure(errors="replace")
except Exception:
    pass

# ------------------------------------------------------------
# Portable paths -- always relative to THIS file, never to the current
# working directory or any one person's machine.
# ------------------------------------------------------------
SCRIPT_DIR = Path(__file__).resolve().parents[2]
CONFIG_DIR = SCRIPT_DIR / "config"
DATA_DIR = SCRIPT_DIR / "data" / "raw"
SECRETS_ENV_PATH = CONFIG_DIR / "secrets.env"
SECRETS_ENV_EXAMPLE_PATH = CONFIG_DIR / "secrets.env.example"

DEFAULT_SHARE_URL = (
    "https://worldhealthorg.sharepoint.com/:x:/r/sites/AF-pep/GISWORKSPACE/"
    "_layouts/15/Doc.aspx?sourcedoc=%7BE89A19B2-233D-4E86-872F-DC8DC2A74F01%7D"
    "&file=AllPolioviruses_20250106.xlsx&action=default&mobileredirect=true"
)
DEFAULT_OUTPUT_NAME = "afro_line_list.xlsx"
GRAPH_ENDPOINT = "https://graph.microsoft.com/v1.0"
MIN_BYTES = 1000  # below this, treat the download as failed/truncated


def log(msg):
    try:
        print(f"[fetch_afro_line_list] {msg}", flush=True)
    except UnicodeEncodeError:
        print(f"[fetch_afro_line_list] {str(msg).encode('ascii', errors='replace').decode('ascii')}", flush=True)


def load_secrets_env(path: Path):
    """Minimal, dependency-free .env loader: reads KEY=VALUE lines from
    `path` (if it exists) into os.environ, without overwriting a variable
    that's already set in the real shell environment (so `setx`/`export`
    always wins over the file). Blank lines and lines starting with '#'
    are ignored. Silently does nothing if the file doesn't exist."""
    if not path.exists():
        return
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, value = line.partition("=")
        key = key.strip()
        value = value.strip()
        if key and key not in os.environ:
            os.environ[key] = value


def ensure_config_bootstrapped() -> bool:
    """First-run-on-a-new-station helper: if config/secrets.env doesn't
    exist yet but the .example template does, copy it into place and stop
    with clear instructions, instead of failing later with a confusing
    'credentials not set' error. Returns True if secrets.env exists (or
    was just created) and the caller can proceed to load it; False if the
    caller should stop and let the person fill it in first."""
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    if SECRETS_ENV_PATH.exists():
        return True

    if SECRETS_ENV_EXAMPLE_PATH.exists():
        shutil.copy2(SECRETS_ENV_EXAMPLE_PATH, SECRETS_ENV_PATH)
        log(f"First run on this station: created {SECRETS_ENV_PATH} from the template.")
        log(f"Open it and fill in SHAREPOINT_TENANT_ID / SHAREPOINT_CLIENT_ID / "
            f"SHAREPOINT_CLIENT_SECRET, then run this script again.")
        return False

    log(f"WARNING: neither {SECRETS_ENV_PATH} nor {SECRETS_ENV_EXAMPLE_PATH} exists. "
        f"Create {SECRETS_ENV_PATH} yourself with SHAREPOINT_TENANT_ID / SHAREPOINT_CLIENT_ID / "
        f"SHAREPOINT_CLIENT_SECRET, or set those three as real environment variables.")
    return True  # nothing to bootstrap from -- let the credential check below give the real error


def _graph_share_id(share_url: str) -> str:
    """Encode a SharePoint/OneDrive sharing URL into the base64 'shares'
    id Microsoft Graph's /shares/{id} endpoint expects (documented Graph
    behavior for resolving any sharing link into a driveItem, regardless
    of which folder the file actually lives in): base64-encode the URL,
    strip '=' padding, make it URL-safe, and prefix with 'u!'."""
    import base64
    b64 = base64.b64encode(share_url.encode("utf-8")).decode("utf-8")
    b64 = b64.rstrip("=").replace("/", "_").replace("+", "-")
    return "u!" + b64


def get_graph_token(tenant_id: str, client_id: str, client_secret: str) -> str:
    """App-only client-credentials token: authenticates as the registered
    application itself, not as a person -- no browser login, no MFA."""
    import requests
    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    data = {
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials",
    }
    r = requests.post(url, data=data, timeout=30)
    r.raise_for_status()
    return r.json()["access_token"]


def download_sharepoint_file(share_url: str, output_path: Path, min_bytes: int = MIN_BYTES) -> bool:
    """Download a single file from a WHO SharePoint sharing link and save
    it at output_path, backing up whatever was already there (into a
    backups/ subfolder next to it). Never raises -- any failure is logged
    and reported back as False."""
    import requests
    import io
    import zipfile
    import tempfile

    tenant_id = os.environ.get("SHAREPOINT_TENANT_ID", "").strip()
    client_id = os.environ.get("SHAREPOINT_CLIENT_ID", "").strip()
    client_secret = os.environ.get("SHAREPOINT_CLIENT_SECRET", "").strip()
    if not all([tenant_id, client_id, client_secret]):
        log(f"FAILED: SHAREPOINT_TENANT_ID / SHAREPOINT_CLIENT_ID / SHAREPOINT_CLIENT_SECRET "
            f"are not all set. Edit {SECRETS_ENV_PATH} and fill them in, or set them as real "
            f"environment variables. Treat them like passwords -- never hardcode them in this "
            f"script or paste them anywhere else.")
        return False

    log("Authenticating to Microsoft Graph...")
    try:
        token = get_graph_token(tenant_id, client_id, client_secret)
    except Exception as e:
        log(f"FAILED: could not authenticate to Microsoft Graph ({type(e).__name__}).")
        return False

    headers = {"Authorization": f"Bearer {token}"}
    share_id = _graph_share_id(share_url)
    try:
        meta = requests.get(f"{GRAPH_ENDPOINT}/shares/{share_id}/driveItem", headers=headers, timeout=60)
        meta.raise_for_status()
        item = meta.json()

        download_url = item.get("@microsoft.graph.downloadUrl")
        if download_url:
            file_resp = requests.get(download_url, timeout=180)
            file_resp.raise_for_status()
            file_bytes = file_resp.content
        else:
            # Fall back to the drive item's own /content endpoint if the
            # metadata response didn't include a direct download URL.
            drive_id = item["parentReference"]["driveId"]
            item_id = item["id"]
            content_resp = requests.get(
                f"{GRAPH_ENDPOINT}/drives/{drive_id}/items/{item_id}/content",
                headers=headers, timeout=180,
            )
            content_resp.raise_for_status()
            file_bytes = content_resp.content
    except Exception as e:
        log(f"FAILED: could not download the file from SharePoint ({type(e).__name__}).")
        return False

    if len(file_bytes) < min_bytes:
        log(f"FAILED: downloaded file looks too small ({len(file_bytes)} bytes) -- "
            f"treating this as a failed download. Nothing was replaced.")
        return False

    try:
        with zipfile.ZipFile(io.BytesIO(file_bytes)) as workbook:
            if not {"[Content_Types].xml", "xl/workbook.xml"}.issubset(workbook.namelist()):
                raise ValueError("Not an Excel workbook")
            if workbook.testzip() is not None:
                raise ValueError("Corrupt workbook")
    except (zipfile.BadZipFile, ValueError):
        log("FAILED: invalid XLSX download. Nothing was replaced.")
        return False

    temporary_path = None
    try:
        if output_path.exists():
            backup_dir = output_path.parent / "backups"
            backup_dir.mkdir(parents=True, exist_ok=True)
            backup_path = backup_dir / f"{output_path.stem}_{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}{output_path.suffix}"
            shutil.copy2(output_path, backup_path)
            log(f"Backed up previous file -> {backup_path}")
        output_path.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.NamedTemporaryFile(dir=output_path.parent, suffix=".tmp", delete=False) as temporary:
            temporary_path = Path(temporary.name)
            temporary.write(file_bytes)
        os.replace(temporary_path, output_path)
    except OSError as e:
        log(f"FAILED: could not write {output_path} ({type(e).__name__}). This almost always means the file "
            f"is currently open in Excel or another program -- close it and re-run.")
        return False

    finally:
        if temporary_path is not None:
            temporary_path.unlink(missing_ok=True)

    log(f"SUCCESS: {output_path} saved ({len(file_bytes)} bytes).")
    return True


def main():
    parser = argparse.ArgumentParser(
        description="Download the WHO AFRO polio line list from SharePoint via Microsoft Graph "
                    "(app-only auth, no browser/MFA) and save it as data/afro_line_list.xlsx. "
                    "All paths are resolved relative to this script's own folder, so the whole "
                    "folder can be copied to any station and just works once config/secrets.env "
                    "is filled in."
    )
    parser.add_argument("--share-url", default=DEFAULT_SHARE_URL,
                        help="SharePoint sharing link to download (default: the AFRO line list link).")
    parser.add_argument("--output-dir", default=str(DATA_DIR),
                        help=f"Directory to save the file into (default: {DATA_DIR}).")
    parser.add_argument("--output-name", default=DEFAULT_OUTPUT_NAME,
                        help=f"Filename to save as (default: {DEFAULT_OUTPUT_NAME}).")
    parser.add_argument("--secrets-env", default=str(SECRETS_ENV_PATH),
                        help=f"Path to a KEY=VALUE file providing SHAREPOINT_TENANT_ID / "
                             f"SHAREPOINT_CLIENT_ID / SHAREPOINT_CLIENT_SECRET if they aren't "
                             f"already set as environment variables (default: {SECRETS_ENV_PATH}).")
    args = parser.parse_args()

    # Only auto-bootstrap when the caller is using the default secrets
    # path -- someone who passed --secrets-env explicitly is pointing at
    # their own file on purpose and shouldn't have a new one created for
    # them next to this script.
    if args.secrets_env == str(SECRETS_ENV_PATH):
        if not ensure_config_bootstrapped():
            sys.exit(1)

    load_secrets_env(Path(args.secrets_env))

    output_path = Path(args.output_dir) / args.output_name
    ok = download_sharepoint_file(args.share_url, output_path)
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
