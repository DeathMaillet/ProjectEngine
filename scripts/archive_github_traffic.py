from __future__ import annotations

import csv
import json
import os
import sys
import urllib.error
import urllib.request
from datetime import datetime, timezone
from pathlib import Path

API_VERSION = "2026-03-10"
repo = os.environ.get("TRAFFIC_SOURCE_REPOSITORY", "").strip()
token = os.environ.get("TRAFFIC_READ_TOKEN", "").strip()
archive_dir = Path(os.environ.get("TRAFFIC_ARCHIVE_DIR", "")).expanduser()

if not repo or "/" not in repo:
    raise SystemExit("TRAFFIC_SOURCE_REPOSITORY is missing or invalid.")
if not token:
    raise SystemExit("TRAFFIC_READ_TOKEN is missing.")
if not str(archive_dir):
    raise SystemExit("TRAFFIC_ARCHIVE_DIR is missing.")

owner, name = repo.split("/", 1)
api_base = f"https://api.github.com/repos/{owner}/{name}/traffic"
headers = {
    "Accept": "application/vnd.github+json",
    "Authorization": f"Bearer {token}",
    "X-GitHub-Api-Version": API_VERSION,
    "User-Agent": "ProjectEngine-private-traffic-archiver",
}

def api_get(path: str):
    request = urllib.request.Request(api_base + path, headers=headers)
    try:
        with urllib.request.urlopen(request, timeout=30) as response:
            return json.load(response)
    except urllib.error.HTTPError as exc:
        body = exc.read().decode("utf-8", errors="replace")
        print(f"GitHub API error {exc.code} for {path}: {body}", file=sys.stderr)
        raise

def load_csv(path: Path) -> list[dict[str, str]]:
    if not path.exists():
        return []
    with path.open("r", newline="", encoding="utf-8") as handle:
        return list(csv.DictReader(handle))

def write_csv(path: Path, fields: list[str], rows: list[dict[str, str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields)
        writer.writeheader()
        writer.writerows(rows)

def upsert_rows(path: Path, fields: list[str], incoming: list[dict], key_fields: list[str]) -> None:
    by_key: dict[tuple[str, ...], dict[str, str]] = {}
    for row in load_csv(path):
        key = tuple(row.get(k, "") for k in key_fields)
        by_key[key] = {field: row.get(field, "") for field in fields}
    for row in incoming:
        normalized = {field: str(row.get(field, "")) for field in fields}
        key = tuple(normalized[k] for k in key_fields)
        by_key[key] = normalized
    rows = list(by_key.values())
    rows.sort(key=lambda row: tuple(row.get(k, "") for k in key_fields))
    write_csv(path, fields, rows)

now = datetime.now(timezone.utc).replace(microsecond=0)
snapshot_utc = now.isoformat()
snapshot_date = now.date().isoformat()

data_dir = archive_dir / "data"
raw_dir = data_dir / "raw"
data_dir.mkdir(parents=True, exist_ok=True)
raw_dir.mkdir(parents=True, exist_ok=True)

views = api_get("/views?per=day")
clones = api_get("/clones?per=day")
referrers = api_get("/popular/referrers")
popular_paths = api_get("/popular/paths")

view_rows = [{"date": i["timestamp"][:10], "views": i["count"], "unique_visitors": i["uniques"]} for i in views.get("views", [])]
upsert_rows(data_dir / "daily_views.csv", ["date", "views", "unique_visitors"], view_rows, ["date"])

clone_rows = [{"date": i["timestamp"][:10], "clones": i["count"], "unique_cloners": i["uniques"]} for i in clones.get("clones", [])]
upsert_rows(data_dir / "daily_clones.csv", ["date", "clones", "unique_cloners"], clone_rows, ["date"])

ref_fields = ["snapshot_date", "referrer", "views", "unique_visitors"]
ref_rows = [{"snapshot_date": snapshot_date, "referrer": i["referrer"], "views": i["count"], "unique_visitors": i["uniques"]} for i in referrers]
existing_refs = [r for r in load_csv(data_dir / "referrers_history.csv") if r.get("snapshot_date") != snapshot_date]
existing_refs.extend({k: str(r.get(k, "")) for k in ref_fields} for r in ref_rows)
existing_refs.sort(key=lambda r: (r.get("snapshot_date", ""), r.get("referrer", "")))
write_csv(data_dir / "referrers_history.csv", ref_fields, existing_refs)

path_fields = ["snapshot_date", "path", "title", "views", "unique_visitors"]
path_rows = [{"snapshot_date": snapshot_date, "path": i["path"], "title": i.get("title", ""), "views": i["count"], "unique_visitors": i["uniques"]} for i in popular_paths]
existing_paths = [r for r in load_csv(data_dir / "popular_paths_history.csv") if r.get("snapshot_date") != snapshot_date]
existing_paths.extend({k: str(r.get(k, "")) for k in path_fields} for r in path_rows)
existing_paths.sort(key=lambda r: (r.get("snapshot_date", ""), r.get("path", "")))
write_csv(data_dir / "popular_paths_history.csv", path_fields, existing_paths)

roll_fields = ["snapshot_date", "snapshot_utc", "rolling_views", "rolling_unique_visitors", "rolling_clones", "rolling_unique_cloners"]
roll_row = {
    "snapshot_date": snapshot_date,
    "snapshot_utc": snapshot_utc,
    "rolling_views": views.get("count", ""),
    "rolling_unique_visitors": views.get("uniques", ""),
    "rolling_clones": clones.get("count", ""),
    "rolling_unique_cloners": clones.get("uniques", ""),
}
rolls = [r for r in load_csv(data_dir / "rolling_14d_snapshots.csv") if r.get("snapshot_date") != snapshot_date]
rolls.append({k: str(roll_row.get(k, "")) for k in roll_fields})
rolls.sort(key=lambda r: r.get("snapshot_date", ""))
write_csv(data_dir / "rolling_14d_snapshots.csv", roll_fields, rolls)

(raw_dir / f"{snapshot_date}.json").write_text(json.dumps({
    "snapshot_utc": snapshot_utc,
    "repository": repo,
    "views": views,
    "clones": clones,
    "referrers": referrers,
    "popular_paths": popular_paths,
}, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")

print(f"Archived private traffic snapshot for {repo}")
print(f"UTC snapshot: {snapshot_utc}")
print(f"View days returned: {len(view_rows)}")
print(f"Clone days returned: {len(clone_rows)}")
print(f"Referrers returned: {len(ref_rows)}")
print(f"Popular paths returned: {len(path_rows)}")
