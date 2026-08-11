from __future__ import annotations

import csv
import json
import os
import sys
import urllib.error
import urllib.request
from collections import defaultdict
from datetime import datetime, timezone
from pathlib import Path

import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from datetime import datetime as dt


API_VERSION = "2026-03-10"

repo = os.environ.get("TRAFFIC_SOURCE_REPOSITORY", "").strip()
token = os.environ.get("TRAFFIC_READ_TOKEN", "").strip()
archive_dir = Path(os.environ.get("TRAFFIC_ARCHIVE_DIR", "")).expanduser()

if not repo or "/" not in repo:
    raise SystemExit("TRAFFIC_SOURCE_REPOSITORY is missing or invalid.")
if not token:
    raise SystemExit("TRAFFIC_READ_TOKEN is missing.")

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


def ints(rows, field):
    return [int(r.get(field, "0") or 0) for r in rows]


def dates(rows, field="date"):
    return [dt.strptime(r[field], "%Y-%m-%d") for r in rows]


def save_line_chart(path: Path, title: str, x, series: list[tuple[str, list[int]]], ylabel: str):
    if not x:
        return
    fig, ax = plt.subplots(figsize=(11, 4.8))
    for label, values in series:
        ax.plot(x, values, marker="o", linewidth=2, label=label)
    ax.set_title(title)
    ax.set_ylabel(ylabel)
    ax.grid(True, alpha=0.25)
    ax.legend()
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%d %b"))
    fig.autofmt_xdate()
    fig.tight_layout()
    path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(path, dpi=160)
    plt.close(fig)


def save_bar_chart(path: Path, title: str, labels: list[str], values: list[int], xlabel: str):
    if not labels:
        return
    labels = labels[:10]
    values = values[:10]
    fig, ax = plt.subplots(figsize=(11, 5))
    y = list(range(len(labels)))
    ax.barh(y, values)
    ax.set_yticks(y, labels)
    ax.invert_yaxis()
    ax.set_title(title)
    ax.set_xlabel(xlabel)
    ax.grid(True, axis="x", alpha=0.25)
    fig.tight_layout()
    path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(path, dpi=160)
    plt.close(fig)


now = datetime.now(timezone.utc).replace(microsecond=0)
snapshot_utc = now.isoformat()
snapshot_date = now.date().isoformat()

data_dir = archive_dir / "data"
raw_dir = data_dir / "raw"
charts_dir = archive_dir / "charts"
data_dir.mkdir(parents=True, exist_ok=True)
raw_dir.mkdir(parents=True, exist_ok=True)
charts_dir.mkdir(parents=True, exist_ok=True)

views = api_get("/views?per=day")
clones = api_get("/clones?per=day")
referrers = api_get("/popular/referrers")
popular_paths = api_get("/popular/paths")

view_rows = [
    {"date": item["timestamp"][:10], "views": item["count"], "unique_visitors": item["uniques"]}
    for item in views.get("views", [])
]
upsert_rows(
    data_dir / "daily_views.csv",
    ["date", "views", "unique_visitors"],
    view_rows,
    ["date"],
)

clone_rows = [
    {"date": item["timestamp"][:10], "clones": item["count"], "unique_cloners": item["uniques"]}
    for item in clones.get("clones", [])
]
upsert_rows(
    data_dir / "daily_clones.csv",
    ["date", "clones", "unique_cloners"],
    clone_rows,
    ["date"],
)

referrer_rows = [
    {
        "snapshot_date": snapshot_date,
        "referrer": item["referrer"],
        "views": item["count"],
        "unique_visitors": item["uniques"],
    }
    for item in referrers
]
old_referrers = [
    row for row in load_csv(data_dir / "referrers_history.csv")
    if row.get("snapshot_date") != snapshot_date
]
old_referrers.extend(
    {k: str(row.get(k, "")) for k in ["snapshot_date", "referrer", "views", "unique_visitors"]}
    for row in referrer_rows
)
old_referrers.sort(key=lambda row: (row.get("snapshot_date", ""), row.get("referrer", "")))
write_csv(
    data_dir / "referrers_history.csv",
    ["snapshot_date", "referrer", "views", "unique_visitors"],
    old_referrers,
)

path_rows = [
    {
        "snapshot_date": snapshot_date,
        "path": item["path"],
        "title": item.get("title", ""),
        "views": item["count"],
        "unique_visitors": item["uniques"],
    }
    for item in popular_paths
]
old_paths = [
    row for row in load_csv(data_dir / "popular_paths_history.csv")
    if row.get("snapshot_date") != snapshot_date
]
old_paths.extend(
    {k: str(row.get(k, "")) for k in ["snapshot_date", "path", "title", "views", "unique_visitors"]}
    for row in path_rows
)
old_paths.sort(key=lambda row: (row.get("snapshot_date", ""), row.get("path", "")))
write_csv(
    data_dir / "popular_paths_history.csv",
    ["snapshot_date", "path", "title", "views", "unique_visitors"],
    old_paths,
)

rollup_fields = [
    "snapshot_date", "snapshot_utc",
    "rolling_views", "rolling_unique_visitors",
    "rolling_clones", "rolling_unique_cloners",
]
rollup_row = {
    "snapshot_date": snapshot_date,
    "snapshot_utc": snapshot_utc,
    "rolling_views": views.get("count", ""),
    "rolling_unique_visitors": views.get("uniques", ""),
    "rolling_clones": clones.get("count", ""),
    "rolling_unique_cloners": clones.get("uniques", ""),
}
rollups = [
    row for row in load_csv(data_dir / "rolling_14d_snapshots.csv")
    if row.get("snapshot_date") != snapshot_date
]
rollups.append({field: str(rollup_row.get(field, "")) for field in rollup_fields})
rollups.sort(key=lambda row: row.get("snapshot_date", ""))
write_csv(data_dir / "rolling_14d_snapshots.csv", rollup_fields, rollups)

(raw_dir / f"{snapshot_date}.json").write_text(
    json.dumps({
        "snapshot_utc": snapshot_utc,
        "repository": repo,
        "views": views,
        "clones": clones,
        "referrers": referrers,
        "popular_paths": popular_paths,
    }, indent=2, ensure_ascii=False) + "\n",
    encoding="utf-8",
)

# ----- charts -----

all_views = load_csv(data_dir / "daily_views.csv")
save_line_chart(
    charts_dir / "daily_views.png",
    "ProjectEngine — daily repository traffic",
    dates(all_views),
    [
        ("Views", ints(all_views, "views")),
        ("Unique visitors", ints(all_views, "unique_visitors")),
    ],
    "Count",
)

all_clones = load_csv(data_dir / "daily_clones.csv")
save_line_chart(
    charts_dir / "daily_clones.png",
    "ProjectEngine — daily clones",
    dates(all_clones),
    [
        ("Clones", ints(all_clones, "clones")),
        ("Unique cloners", ints(all_clones, "unique_cloners")),
    ],
    "Count",
)

all_rollups = load_csv(data_dir / "rolling_14d_snapshots.csv")
if len(all_rollups) >= 1:
    save_line_chart(
        charts_dir / "rolling_14d.png",
        "ProjectEngine — rolling 14-day traffic",
        [dt.strptime(r["snapshot_date"], "%Y-%m-%d") for r in all_rollups],
        [
            ("Views", ints(all_rollups, "rolling_views")),
            ("Unique visitors", ints(all_rollups, "rolling_unique_visitors")),
            ("Clones", ints(all_rollups, "rolling_clones")),
            ("Unique cloners", ints(all_rollups, "rolling_unique_cloners")),
        ],
        "Count",
    )

latest_refs = sorted(referrers, key=lambda x: x["uniques"], reverse=True)
save_bar_chart(
    charts_dir / "latest_referrers.png",
    f"Top referrers — snapshot {snapshot_date}",
    [x["referrer"] for x in latest_refs],
    [x["uniques"] for x in latest_refs],
    "Unique visitors in GitHub rolling window",
)

latest_paths = sorted(popular_paths, key=lambda x: x["uniques"], reverse=True)
save_bar_chart(
    charts_dir / "latest_paths.png",
    f"Most visited repository content — snapshot {snapshot_date}",
    [x.get("title") or x["path"] for x in latest_paths],
    [x["uniques"] for x in latest_paths],
    "Unique visitors in GitHub rolling window",
)

# Per-source history: the values are rolling-window snapshots, not exact daily referrals.
hist_refs = load_csv(data_dir / "referrers_history.csv")
sources = defaultdict(list)
all_snapshot_dates = sorted({r["snapshot_date"] for r in hist_refs})
for source in sorted({r["referrer"] for r in hist_refs}):
    by_date = {r["snapshot_date"]: int(r["unique_visitors"] or 0)
               for r in hist_refs if r["referrer"] == source}
    values = [by_date.get(d, 0) for d in all_snapshot_dates]
    sources[source] = values

if all_snapshot_dates:
    ranked_sources = sorted(
        sources,
        key=lambda s: max(sources[s]) if sources[s] else 0,
        reverse=True
    )[:8]
    save_line_chart(
        charts_dir / "referrers_history.png",
        "Referrer evolution — rolling-window snapshots",
        [dt.strptime(d, "%Y-%m-%d") for d in all_snapshot_dates],
        [(s, sources[s]) for s in ranked_sources],
        "Unique visitors in GitHub rolling window",
    )

# ----- private README dashboard -----

latest_views = int(views.get("count", 0) or 0)
latest_unique = int(views.get("uniques", 0) or 0)
latest_clones = int(clones.get("count", 0) or 0)
latest_unique_cloners = int(clones.get("uniques", 0) or 0)

readme = f"""# ProjectEngine Traffic Archive

Private long-term traffic archive for `TMailletFR/ProjectEngine`.

_Last updated: {snapshot_utc}_

## Current GitHub rolling window

| Metric | Value |
|---|---:|
| Views | **{latest_views}** |
| Unique visitors | **{latest_unique}** |
| Clones | **{latest_clones}** |
| Unique cloners | **{latest_unique_cloners}** |

## Daily traffic

![Daily views](charts/daily_views.png)

## Daily clones

![Daily clones](charts/daily_clones.png)

## Rolling 14-day trend

![Rolling 14-day trend](charts/rolling_14d.png)

## Current referrers

![Latest referrers](charts/latest_referrers.png)

## Referrer evolution

![Referrer history](charts/referrers_history.png)

> Referrer values are GitHub rolling-window snapshots. They are not exact per-day attribution until enough history exists to reconstruct what entered and left the rolling window.

## Most visited repository content

![Popular paths](charts/latest_paths.png)

## Raw data

The permanent datasets are stored under [`data/`](data/), including daily views, daily clones, referrer snapshots, popular-path snapshots and raw API responses.
"""

(archive_dir / "README.md").write_text(readme, encoding="utf-8")

print(f"Archived and rendered private dashboard for {repo} at {snapshot_utc}")
