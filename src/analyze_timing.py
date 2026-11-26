"""공시가 한국 업무시간 밖에 얼마나 몰리는지 집계한다 (기획서 3쪽 숫자).

입력: data/filings_all.csv, data/8k_list.csv (collect_filings.py 결과)
업무시간 = 한국 시간 평일 9~18시. SEC 접수 시각(UTC)을 한국 시간(UTC+9)으로 바꿔 판단한다.
"""
import csv
from collections import Counter
from datetime import datetime, timedelta
from pathlib import Path

DATA = Path(__file__).resolve().parent.parent / "data"
KST = timedelta(hours=9)


def load(name):
    with open(DATA / name, encoding="utf-8-sig") as fp:
        return list(csv.DictReader(fp))


def to_kst(utc_text):
    return datetime.strptime(utc_text[:19], "%Y-%m-%dT%H:%M:%S") + KST


def main():
    allf, eight_k = load("filings_all.csv"), load("8k_list.csv")
    forms = Counter(r["form"] for r in allf)
    print(f"전체 공시 {len(allf):,}건 중 8-K {len(eight_k)}건 ({len(eight_k) / len(allf):.1%})")
    for form, n in forms.most_common(3):
        print(f"  {form:6s} {n:,}건 ({n / len(allf):.0%})")

    times = [to_kst(r["acceptanceDateTime"]) for r in eight_k]
    in_hours = [t for t in times if t.weekday() < 5 and 9 <= t.hour < 18]
    out = len(times) - len(in_hours)
    print(f"8-K {len(times)}건 중 한국 업무시간 밖 {out}건 ({out / len(times):.1%})")
    print(f"  05~07시(05:00~06:59) 접수 {sum(5 <= t.hour < 7 for t in times)}건")

    by_hour = Counter(t.hour for t in times)
    print("시간대별(한국 시간):", " ".join(f"{h}시:{by_hour[h]}" for h in range(24) if by_hour[h]))


if __name__ == "__main__":
    main()
