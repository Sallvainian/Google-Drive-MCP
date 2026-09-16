import json, sys, tiktoken

if len(sys.argv) < 2:
    sys.stderr.write("usage: measure-token-cost.py <tools-list.json> [report.json]\n")
    sys.exit(2)

enc = tiktoken.get_encoding("cl100k_base")
d = json.load(open(sys.argv[1]))
tools = d["tools"]
rows = []
for t in tools:
    s = json.dumps(t, separators=(",", ":"))
    rows.append((t["name"], len(enc.encode(s)),
                 len(enc.encode(t.get("description") or "")),
                 len(enc.encode(json.dumps(t.get("inputSchema", {}), separators=(",", ":"))))))
rows.sort(key=lambda r: -r[1])
total = sum(r[1] for r in rows)
n = len(rows)
print(f"TOOLS: {n}")
print(f"TOTAL TOKENS (cl100k_base, compact JSON): {total:,}")
print(f"AVERAGE: {total/n:.0f}")
med = sorted(r[1] for r in rows)[n//2]
print(f"MEDIAN: {med}")
print(f"MAX: {rows[0][0]} = {rows[0][1]:,} ({rows[0][1]/total*100:.1f}%)")
print(f"MIN: {rows[-1][0]} = {rows[-1][1]}")
desc = sum(r[2] for r in rows); sch = sum(r[3] for r in rows)
print(f"SPLIT: descriptions {desc:,} ({desc/total*100:.0f}%) | inputSchema {sch:,} ({sch/total*100:.0f}%)")
print(f"\nTOP 10 BY COST:")
for name, tot, dsc, sc in rows[:10]:
    print(f"  {tot:6,}  {name}  (desc {dsc}, schema {sc})")
print(f"\nOVER SENTRY'S 1000-TOKEN INVESTIGATE THRESHOLD: {sum(1 for r in rows if r[1] > 1000)} tools")
print(f"OVER SENTRY'S 500-TOKEN NEW-TOOL BUDGET: {sum(1 for r in rows if r[1] > 500)} tools")

if len(sys.argv) >= 3:
    report = {
        "count": n,
        "total": total,
        "average": int(f"{total/n:.0f}") if n else 0,
        "tools": [{"name": name, "tokens": tot} for name, tot, dsc, sc in rows],
    }
    with open(sys.argv[2], "w") as fh:
        json.dump(report, fh, indent=2)
        fh.write("\n")
