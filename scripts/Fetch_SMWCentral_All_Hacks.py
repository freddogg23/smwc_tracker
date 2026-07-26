#!/usr/bin/env python3
from __future__ import annotations
import csv, json, re, sys, time
from collections import Counter
from datetime import datetime
from pathlib import Path
from typing import Any
from urllib.parse import urlencode
from urllib.request import Request, urlopen

API_URL = "https://www.smwcentral.net/ajax.php"
DETAIL_URL = "https://www.smwcentral.net/?a=details&id={id}&p=section"
DIFFICULTY = {"diff_1":"Newcomer","diff_2":"Casual","diff_3":"Intermediate","diff_4":"Advanced","diff_5":"Expert","diff_6":"Master","diff_7":"Grandmaster"}

def text(v: Any) -> str:
    if v is None: return ""
    if isinstance(v, (str,int,float)): return str(v).strip()
    if isinstance(v, dict):
        for k in ("name","username","display_name","label","title","text","value"):
            if str(v.get(k) or "").strip(): return str(v[k]).strip()
    return str(v).strip()

def joined(v: Any) -> str:
    return ", ".join(filter(None, (text(x) for x in v))) if isinstance(v, list) else text(v)

def type_text(v: Any) -> str:
    vals = v if isinstance(v, list) else re.split(r"[,;]", str(v or ""))
    mapping={"standard":"Standard","kaizo":"Kaizo","puzzle":"Puzzle","tool_assisted":"Tool-Assisted","toolassisted":"Tool-Assisted","pit":"Pit","troll":"Troll"}
    out=[]
    for x in vals:
        t=text(x)
        if not t: continue
        key=re.sub(r"[- ]","_",t.lower())
        display=mapping.get(key,t.replace("_"," ").title())
        if display not in out: out.append(display)
    return ", ".join(out)

def normalize_url(v: Any) -> str:
    u=str(v or "").strip()
    if u.startswith("//"): return "https:"+u
    if u.startswith("/"): return "https://www.smwcentral.net"+u
    return u

def fetch(page:int) -> dict[str,Any]:
    q=urlencode({"a":"getsectionlist","s":"smwhacks","n":page,"u":"0"})
    req=Request(f"{API_URL}?{q}",headers={"User-Agent":"SMW-Hack-Tracker-Updater/1.0","Accept":"application/json,text/plain,*/*","Referer":"https://www.smwcentral.net/?p=section&s=smwhacks"})
    last=None
    for attempt in range(4):
        try:
            with urlopen(req,timeout=60) as r: return json.loads(r.read().decode("utf-8"))
        except Exception as e:
            last=e
            if attempt<3: time.sleep(2**(attempt+1))
    raise RuntimeError(f"Page {page} failed: {last}")

def main() -> int:
    root=Path(__file__).resolve().parents[1]
    outdir=root/"data"; outdir.mkdir(exist_ok=True)
    rows_by_id={}; page=1; last_page=1; reported=0
    while page<=last_page:
        payload=fetch(page); last_page=int(payload.get("last_page") or 1); reported=int(payload.get("total") or 0)
        for hack in payload.get("data") or []:
            hid=str(hack.get("id") or "").strip(); title=str(hack.get("name") or "").strip()
            if not hid or not title or hid in rows_by_id: continue
            raw=hack.get("raw_fields") or {}
            m=re.search(r"-?\d+",str(raw.get("length",hack.get("length",0))))
            diff=joined(raw.get("difficulty",hack.get("difficulty")))
            ts=hack.get("time")
            try: added=datetime.fromtimestamp(int(ts)).strftime("%Y-%m-%d") if ts else ""
            except Exception: added=""
            rows_by_id[hid]={"ROM Hack Title":title,"Created By":joined(hack.get("authors")) or "Unknown","Exits":int(m.group()) if m else 0,"Difficulty":DIFFICULTY.get(diff,diff.removeprefix("diff_").replace("_"," ").title() if diff else "Unranked"),"Type":type_text(raw.get("type",hack.get("type"))),"Added Date":added,"SMWC ID":hid,"SMWCentral Page URL":DETAIL_URL.format(id=hid),"Direct Download URL":normalize_url(hack.get("download_url"))}
        print(f"Fetched {page}/{last_page}: {len(rows_by_id)}/{reported or '?'}",flush=True); page+=1
        if page<=last_page: time.sleep(2.5)
    ordered=sorted(rows_by_id.values(),key=lambda r:(r["ROM Hack Title"].casefold(),int(r["SMWC ID"])))
    counts=Counter(r["ROM Hack Title"].casefold() for r in ordered)
    rows=[]
    for r in ordered:
        title=r["ROM Hack Title"]; label=title if counts[title.casefold()]==1 else f'{title} [SMWC #{r["SMWC ID"]}]'
        rows.append({"Dropdown Selection":label,**r})
    headers=["Dropdown Selection","ROM Hack Title","Created By","Exits","Difficulty","Type","Added Date","SMWC ID","SMWCentral Page URL","Direct Download URL"]
    with (outdir/"SMWCentral_All_Moderated_Hacks.csv").open("w",newline="",encoding="utf-8-sig") as f:
        w=csv.DictWriter(f,fieldnames=headers); w.writeheader(); w.writerows(rows)
    (outdir/"SMWCentral_All_Moderated_Hacks.json").write_text(json.dumps(rows,ensure_ascii=False,indent=2),encoding="utf-8")
    print(f"Saved {len(rows)} hacks.")
    return 2 if reported and len(rows)!=reported else 0

if __name__=="__main__":
    try: raise SystemExit(main())
    except KeyboardInterrupt: raise SystemExit(130)
    except Exception as exc: print(f"ERROR: {exc}",file=sys.stderr); raise SystemExit(1)
