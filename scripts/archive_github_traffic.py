import csv, json, os, urllib.request
from datetime import datetime, timezone
from pathlib import Path

OUT = Path("analytics/github-traffic")
repo = os.environ["GITHUB_REPOSITORY"]
token = os.environ["GITHUB_TOKEN"]
owner, name = repo.split("/", 1)
base = f"https://api.github.com/repos/{owner}/{name}/traffic"
headers = {
    "Accept": "application/vnd.github+json",
    "Authorization": f"Bearer {token}",
    "X-GitHub-Api-Version": "2022-11-28",
    "User-Agent": "ProjectEngine-traffic-archiver",
}

def get(path):
    req = urllib.request.Request(base + path, headers=headers)
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.load(r)

def merge_csv(path, fields, rows, keys):
    existing, seen = [], set()
    if path.exists():
        with path.open(newline="", encoding="utf-8") as f:
            for row in csv.DictReader(f):
                existing.append(row)
                seen.add(tuple(row[k] for k in keys))
    for row in rows:
        row = {k: str(row.get(k, "")) for k in fields}
        key = tuple(row[k] for k in keys)
        if key not in seen:
            existing.append(row); seen.add(key)
    existing.sort(key=lambda r: tuple(r[k] for k in keys))
    with path.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader(); w.writerows(existing)

OUT.mkdir(parents=True, exist_ok=True)
now = datetime.now(timezone.utc).replace(microsecond=0)
day = now.date().isoformat()

views = get("/views?per=day")
clones = get("/clones?per=day")
refs = get("/popular/referrers")
paths = get("/popular/paths")

merge_csv(OUT/"daily_views.csv",
          ["date","views","unique_visitors"],
          [{"date":x["timestamp"][:10],"views":x["count"],"unique_visitors":x["uniques"]} for x in views.get("views",[])],
          ["date"])

merge_csv(OUT/"daily_clones.csv",
          ["date","clones","unique_cloners"],
          [{"date":x["timestamp"][:10],"clones":x["count"],"unique_cloners":x["uniques"]} for x in clones.get("clones",[])],
          ["date"])

merge_csv(OUT/"referrers_history.csv",
          ["snapshot_date","referrer","views","unique_visitors"],
          [{"snapshot_date":day,"referrer":x["referrer"],"views":x["count"],"unique_visitors":x["uniques"]} for x in refs],
          ["snapshot_date","referrer"])

merge_csv(OUT/"popular_paths_history.csv",
          ["snapshot_date","path","title","views","unique_visitors"],
          [{"snapshot_date":day,"path":x["path"],"title":x.get("title",""),"views":x["count"],"unique_visitors":x["uniques"]} for x in paths],
          ["snapshot_date","path"])

raw = OUT/"raw"
raw.mkdir(exist_ok=True)
(raw/f"{day}.json").write_text(json.dumps({
    "snapshot_utc": now.isoformat(),
    "repository": repo,
    "views": views,
    "clones": clones,
    "referrers": refs,
    "popular_paths": paths
}, indent=2), encoding="utf-8")

print(f"Archived {repo} traffic for {day}")
