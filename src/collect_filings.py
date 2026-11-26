"""KIC 13F 상위 15종목의 1년치 SEC 공시 목록을 EDGAR Submissions API로 수집한다.

출력
  data/filings_all.csv : 기간 내 전체 공시 (종목, 서식, 접수일, 접수 시각 UTC, 8-K 항목)
  data/8k_list.csv     : 그중 주요 사건 공시(8-K)만

SEC 공정 접근 정책에 따라 User-Agent를 밝히고 초당 10회 미만으로 호출한다.
  SEC_USER_AGENT="이름 연락처" python src/collect_filings.py
"""
import csv
import json
import os
import time
import urllib.request
from pathlib import Path

TICKERS = ["NVDA", "AAPL", "GOOGL", "MSFT", "AMZN", "AVGO", "MU", "META",
           "TSLA", "LLY", "AMD", "JPM", "JNJ", "BRK.B", "INTC"]
START, END = "2025-11-06", "2026-11-06"          # 접수일 기준 [START, END) 1년
UA = os.environ.get("SEC_USER_AGENT", "sentinel-research-bot research@example.com")
OUT = Path(__file__).resolve().parent.parent / "data"


def get_json(url):
    req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept-Encoding": "identity"})
    with urllib.request.urlopen(req, timeout=60) as r:
        data = json.load(r)
    time.sleep(0.15)
    return data


def cik_map():
    rows = get_json("https://www.sec.gov/files/company_tickers.json").values()
    m = {r["ticker"]: r["cik_str"] for r in rows}
    return {t: m[t.replace(".", "-")] for t in TICKERS}


def filings_for(cik):
    """recent 블록과 과거 페이지(files)를 합쳐 기간 안의 공시만 돌려준다."""
    sub = get_json(f"https://data.sec.gov/submissions/CIK{cik:010d}.json")
    blocks = [sub["filings"]["recent"]]
    for f in sub["filings"].get("files", []):
        if f["filingTo"] >= START:                 # 기간과 겹치는 페이지만
            blocks.append(get_json(f"https://data.sec.gov/submissions/{f['name']}"))
    for b in blocks:
        for i in range(len(b["accessionNumber"])):
            if START <= b["filingDate"][i] < END:
                yield {k: b[k][i] for k in ("form", "filingDate", "acceptanceDateTime",
                                            "items", "accessionNumber", "primaryDocument")}


def main():
    OUT.mkdir(exist_ok=True)
    rows, seen = [], set()
    for t, cik in cik_map().items():
        for f in filings_for(cik):
            if f["accessionNumber"] in seen:
                continue
            seen.add(f["accessionNumber"])
            rows.append({"ticker": t, **f})
        print(f"{t:6s} 누적 {len(rows):,}건")

    cols = ["ticker", "form", "filingDate", "acceptanceDateTime", "items", "accessionNumber", "primaryDocument"]
    with open(OUT / "filings_all.csv", "w", newline="", encoding="utf-8") as fp:
        w = csv.DictWriter(fp, cols); w.writeheader(); w.writerows(rows)
    eight_k = [r for r in rows if r["form"] == "8-K"]
    with open(OUT / "8k_list.csv", "w", newline="", encoding="utf-8") as fp:
        w = csv.DictWriter(fp, cols); w.writeheader(); w.writerows(eight_k)
    print(f"전체 {len(rows):,}건, 8-K {len(eight_k)}건 저장")


if __name__ == "__main__":
    main()
