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
    with path.open("r", newline="", encoding="utf-8-sig") as handle:
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


def as_int(value) -> int:
    try:
        return int(str(value).strip())
    except Exception:
        return 0


def parse_dates(rows, field="date"):
    return [dt.strptime(r[field], "%Y-%m-%d") for r in rows]


def ints(rows, field):
    return [as_int(r.get(field, 0)) for r in rows]


def save_line_chart(path: Path, title: str, x, series: list[tuple[str, list[int]]], ylabel: str):
    if not x:
        return
    fig, ax = plt.subplots(figsize=(12, 5.2))
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


def merge_daily_history(api_rows, historical_rows, api_fields, hist_fields):
    """
    Merge historical reconstructed daily rows with canonical API rows.
    API wins on overlapping dates.
    """
    merged = {}
    for r in historical_rows:
        d = r.get("date", "")
        if not d:
            continue
        merged[d] = {
            "date": d,
            api_fields[0]: r.get(hist_fields[0], "0"),
            api_fields[1]: r.get(hist_fields[1], "0"),
            "provenance": r.get("provenance", "screenshot_reconstructed"),
        }
    for r in api_rows:
        d = r.get("date", "")
        if not d:
            continue
        merged[d] = {
            "date": d,
            api_fields[0]: r.get(api_fields[0], "0"),
            api_fields[1]: r.get(api_fields[1], "0"),
            "provenance": "api_exact",
        }
    return [merged[d] for d in sorted(merged)]


def combine_snapshot_history(current_rows, historical_rows, key_fields, value_fields):
    """
    Historical rows are loaded first; canonical automatic rows overwrite exact key collisions.
    """
    merged = {}
    for r in historical_rows:
        key = tuple(r.get(k, "") for k in key_fields)
        merged[key] = {k: r.get(k, "") for k in key_fields + value_fields}
    for r in current_rows:
        key = tuple(r.get(k, "") for k in key_fields)
        merged[key] = {k: r.get(k, "") for k in key_fields + value_fields}
    out = list(merged.values())
    out.sort(key=lambda r: tuple(r.get(k, "") for k in key_fields))
    return out


def load_events(path: Path) -> list[dict[str, str]]:
    rows = load_csv(path)
    rows.sort(key=lambda r: (r.get("date", ""), r.get("category", ""), r.get("label", "")))
    return rows


def event_date_map(events: list[dict[str, str]]) -> dict[str, list[dict[str, str]]]:
    out = defaultdict(list)
    for event in events:
        if event.get("date"):
            out[event["date"]].append(event)
    return out


def save_line_chart_with_events(
    path: Path,
    title: str,
    x,
    series: list[tuple[str, list[int]]],
    ylabel: str,
    events: list[dict[str, str]],
):
    if not x:
        return
    fig, ax = plt.subplots(figsize=(14, 6))
    for label, values in series:
        ax.plot(x, values, marker="o", linewidth=2, label=label)

    ymin, ymax = ax.get_ylim()
    span = ymax - ymin if ymax > ymin else 1

    plotted_dates = set()
    for event in events:
        if str(event.get("plot", "1")).strip() not in ("1", "true", "True", "yes", "YES"):
            continue
        try:
            ed = dt.strptime(event["date"], "%Y-%m-%d")
        except Exception:
            continue
        if ed < min(x) or ed > max(x):
            continue

        ax.axvline(ed, linestyle="--", linewidth=1, alpha=0.45)

        # Stagger labels when several events share nearby dates.
        key = event["date"]
        n = sum(1 for d in plotted_dates if abs((dt.strptime(d, "%Y-%m-%d") - ed).days) <= 1)
        y = ymax - span * (0.05 + 0.08 * (n % 4))
        short = event.get("short_label") or event.get("label") or event.get("category") or "Event"
        ax.text(
            ed, y, short,
            rotation=90,
            va="top",
            ha="right",
            fontsize=8,
            alpha=0.85,
        )
        plotted_dates.add(key)

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

now = datetime.now(timezone.utc).replace(microsecond=0)
snapshot_utc = now.isoformat()
snapshot_date = now.date().isoformat()

data_dir = archive_dir / "data"
historical_dir = data_dir / "historical"
derived_dir = data_dir / "derived"
raw_dir = data_dir / "raw"
charts_dir = archive_dir / "charts"

for d in (data_dir, historical_dir, derived_dir, raw_dir, charts_dir):
    d.mkdir(parents=True, exist_ok=True)

events = load_events(data_dir / "events.csv")

# ----------------------------------------------------------------------
# 1) Fetch canonical current GitHub traffic
# ----------------------------------------------------------------------

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

# ----------------------------------------------------------------------
# 2) Build combined long-term datasets = recovered screenshots + API archive
# ----------------------------------------------------------------------

hist_daily_views = load_csv(historical_dir / "daily_views_reconstructed.csv")
hist_daily_clones = load_csv(historical_dir / "daily_clones_reconstructed.csv")

combined_views = merge_daily_history(
    load_csv(data_dir / "daily_views.csv"),
    hist_daily_views,
    ("views", "unique_visitors"),
    ("views", "unique_visitors"),
)
write_csv(
    derived_dir / "combined_daily_views.csv",
    ["date", "views", "unique_visitors", "provenance"],
    combined_views,
)

combined_clones = merge_daily_history(
    load_csv(data_dir / "daily_clones.csv"),
    hist_daily_clones,
    ("clones", "unique_cloners"),
    ("clones", "unique_cloners"),
)
write_csv(
    derived_dir / "combined_daily_clones.csv",
    ["date", "clones", "unique_cloners", "provenance"],
    combined_clones,
)

hist_refs = load_csv(historical_dir / "referrers_snapshots.csv")
hist_refs_normalized = [
    {
        "snapshot_date": r.get("snapshot_date", ""),
        "referrer": r.get("referrer", ""),
        "views": r.get("views", "0"),
        "unique_visitors": r.get("unique_visitors", "0"),
    }
    for r in hist_refs
]
combined_refs = combine_snapshot_history(
    load_csv(data_dir / "referrers_history.csv"),
    hist_refs_normalized,
    ["snapshot_date", "referrer"],
    ["views", "unique_visitors"],
)
write_csv(
    derived_dir / "combined_referrers_history.csv",
    ["snapshot_date", "referrer", "views", "unique_visitors"],
    combined_refs,
)

hist_rollups = load_csv(historical_dir / "rolling_14d_snapshots.csv")
hist_rollups_normalized = [
    {
        "snapshot_date": r.get("snapshot_date", ""),
        "snapshot_utc": "",
        "rolling_views": r.get("rolling_views", ""),
        "rolling_unique_visitors": r.get("rolling_unique_visitors", ""),
        "rolling_clones": r.get("rolling_clones", ""),
        "rolling_unique_cloners": r.get("rolling_unique_cloners", ""),
    }
    for r in hist_rollups
]
combined_rollups = combine_snapshot_history(
    load_csv(data_dir / "rolling_14d_snapshots.csv"),
    hist_rollups_normalized,
    ["snapshot_date"],
    ["snapshot_utc", "rolling_views", "rolling_unique_visitors", "rolling_clones", "rolling_unique_cloners"],
)
write_csv(
    derived_dir / "combined_rolling_14d_snapshots.csv",
    rollup_fields,
    combined_rollups,
)

# ----------------------------------------------------------------------
# 3) Charts
# ----------------------------------------------------------------------

save_line_chart_with_events(
    charts_dir / "daily_views.png",
    "ProjectEngine — daily repository traffic (historical + API)",
    parse_dates(combined_views),
    [
        ("Views", ints(combined_views, "views")),
        ("Unique visitors", ints(combined_views, "unique_visitors")),
    ],
    "Count",
    events,
)

save_line_chart_with_events(
    charts_dir / "daily_clones.png",
    "ProjectEngine — daily clones (historical + API)",
    parse_dates(combined_clones),
    [
        ("Clones", ints(combined_clones, "clones")),
        ("Unique cloners", ints(combined_clones, "unique_cloners")),
    ],
    "Count",
    events,
)

valid_rollups = [r for r in combined_rollups if r.get("rolling_views", "") != ""]
if valid_rollups:
    save_line_chart(
        charts_dir / "rolling_14d.png",
        "ProjectEngine — rolling 14-day traffic",
        [dt.strptime(r["snapshot_date"], "%Y-%m-%d") for r in valid_rollups],
        [
            ("Views", ints(valid_rollups, "rolling_views")),
            ("Unique visitors", ints(valid_rollups, "rolling_unique_visitors")),
            ("Clones", ints(valid_rollups, "rolling_clones")),
            ("Unique cloners", ints(valid_rollups, "rolling_unique_cloners")),
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

# Referrer evolution across all recovered + current snapshots
sources = defaultdict(dict)
all_snapshot_dates = sorted({r["snapshot_date"] for r in combined_refs if r.get("snapshot_date")})
for r in combined_refs:
    source = r.get("referrer", "")
    if source:
        sources[source][r["snapshot_date"]] = as_int(r.get("unique_visitors", 0))

if all_snapshot_dates:
    ranked_sources = sorted(
        sources,
        key=lambda s: max(sources[s].values()) if sources[s] else 0,
        reverse=True
    )[:8]

    # Missing source on a GitHub top-referrers table is represented as 0 for visualization.
    save_line_chart(
        charts_dir / "referrers_history.png",
        "Referrer evolution — recovered + API rolling snapshots",
        [dt.strptime(d, "%Y-%m-%d") for d in all_snapshot_dates],
        [
            (s, [sources[s].get(d, 0) for d in all_snapshot_dates])
            for s in ranked_sources
        ],
        "Unique visitors in GitHub rolling window",
    )

# ----------------------------------------------------------------------
# 4) Metadata based on all known exact referrer snapshots
# ----------------------------------------------------------------------

metadata_first_seen = []
metadata_peaks = []
metadata_deltas = []

for source in sorted(sources):
    rows = [
        r for r in combined_refs
        if r.get("referrer") == source
    ]
    rows.sort(key=lambda r: r.get("snapshot_date", ""))
    if not rows:
        continue

    first = rows[0]
    metadata_first_seen.append({
        "referrer": source,
        "first_seen_snapshot": first["snapshot_date"],
        "rolling_views": first.get("views", "0"),
        "rolling_unique_visitors": first.get("unique_visitors", "0"),
    })

    peak_v = max(rows, key=lambda r: as_int(r.get("views")))
    peak_u = max(rows, key=lambda r: as_int(r.get("unique_visitors")))
    metadata_peaks.append({
        "referrer": source,
        "peak_rolling_views": peak_v.get("views", "0"),
        "peak_views_snapshot": peak_v.get("snapshot_date", ""),
        "peak_rolling_unique_visitors": peak_u.get("unique_visitors", "0"),
        "peak_uniques_snapshot": peak_u.get("snapshot_date", ""),
    })

    prev = None
    for r in rows:
        if prev is not None:
            metadata_deltas.append({
                "snapshot_date": r.get("snapshot_date", ""),
                "referrer": source,
                "rolling_views": r.get("views", "0"),
                "rolling_unique_visitors": r.get("unique_visitors", "0"),
                "delta_views_vs_previous_snapshot": as_int(r.get("views")) - as_int(prev.get("views")),
                "delta_uniques_vs_previous_snapshot": as_int(r.get("unique_visitors")) - as_int(prev.get("unique_visitors")),
                "previous_snapshot_date": prev.get("snapshot_date", ""),
                "note": "Net change in rolling GitHub snapshot; not exact daily referrals.",
            })
        prev = r

write_csv(
    derived_dir / "referrer_first_seen.csv",
    ["referrer", "first_seen_snapshot", "rolling_views", "rolling_unique_visitors"],
    metadata_first_seen,
)
write_csv(
    derived_dir / "referrer_peaks.csv",
    ["referrer", "peak_rolling_views", "peak_views_snapshot", "peak_rolling_unique_visitors", "peak_uniques_snapshot"],
    metadata_peaks,
)
write_csv(
    derived_dir / "referrer_snapshot_deltas.csv",
    [
        "snapshot_date", "referrer", "rolling_views", "rolling_unique_visitors",
        "delta_views_vs_previous_snapshot", "delta_uniques_vs_previous_snapshot",
        "previous_snapshot_date", "note"
    ],
    metadata_deltas,
)

# ----------------------------------------------------------------------
# 5) Private README dashboard
# ----------------------------------------------------------------------

latest_views = as_int(views.get("count", 0))
latest_unique = as_int(views.get("uniques", 0))
latest_clones = as_int(clones.get("count", 0))
latest_unique_cloners = as_int(clones.get("uniques", 0))

first_daily_date = combined_views[0]["date"] if combined_views else "n/a"
last_daily_date = combined_views[-1]["date"] if combined_views else "n/a"

readme = f"""# ProjectEngine Traffic Archive

Private long-term GitHub traffic archive for `TMailletFR/ProjectEngine`.

_Last updated: {snapshot_utc}_

## Current GitHub rolling window

| Metric | Value |
|---|---:|
| Views | **{latest_views}** |
| Unique visitors | **{latest_unique}** |
| Clones | **{latest_clones}** |
| Unique cloners | **{latest_unique_cloners}** |

## Long-term daily traffic

Recovered historical screenshots are merged with the automatic API archive for the charts below.
On overlapping dates, API data always has priority.

Coverage currently starts on **{first_daily_date}** and runs through **{last_daily_date}** where data is available.

![Daily views](charts/daily_views.png)

## Long-term daily clones

![Daily clones](charts/daily_clones.png)

## Rolling 14-day trend

This chart now includes recovered historical GitHub snapshots as well as the automatic archive.

![Rolling 14-day trend](charts/rolling_14d.png)

## Current referrers

![Latest referrers](charts/latest_referrers.png)

## Referrer evolution

This graph combines exact values transcribed from old GitHub Traffic tables with the daily automatic snapshots.

![Referrer history](charts/referrers_history.png)

> Referrer values are GitHub rolling-window snapshots. A change between two points is a net change of the 14-day window, not exact per-day attribution.

## Most visited repository content

![Popular paths](charts/latest_paths.png)

## Communication & product events

The repository also contains [`data/events.csv`](data/events.csv), our communication / release / infrastructure timeline.

Events marked `plot=1` are drawn directly on the daily traffic and clone charts so traffic spikes can be compared with releases and communication actions.

## Data layout

- [`data/daily_views.csv`](data/daily_views.csv): canonical automatic API archive
- [`data/daily_clones.csv`](data/daily_clones.csv): canonical automatic API archive
- [`data/referrers_history.csv`](data/referrers_history.csv): canonical automatic referrer snapshots
- [`data/historical/`](data/historical/): manually recovered historical screenshots
- [`data/derived/`](data/derived/): merged long-term datasets and metadata generated automatically

The historical reconstruction is kept separate from canonical API data so its provenance remains explicit.
"""

(archive_dir / "README.md").write_text(readme, encoding="utf-8")

print(f"Archived ProjectEngine traffic and rendered merged historical dashboard at {snapshot_utc}")
