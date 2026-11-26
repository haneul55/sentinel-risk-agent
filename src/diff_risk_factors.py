"""연간 보고서(10-K) 위험 요인(Item 1A)에서 새로 생긴 문장을 찾는다 (기획서 4쪽 '조짐 감지').

data/10k/<티커>_<제출일>.txt 두 해를 비교해, 작년 어떤 문장과도
단어 Jaccard 유사도가 0.5 미만인 올해 문장을 '새 문장'으로 본다.
  python src/diff_risk_factors.py            # 전 종목
  python src/diff_risk_factors.py AAPL       # 한 종목
"""
import re
import sys
from collections import defaultdict
from pathlib import Path

DATA = Path(__file__).resolve().parent.parent / "data"
THRESHOLD = 0.5
ITEM_1A = re.compile(r"Item\s*1A\.?\s*Risk\s*Factors", re.I)
ITEM_1B = re.compile(r"Item\s*1B\.?\s*Unresolved", re.I)


def risk_section(text):
    """각 'Item 1B' 바로 앞의 'Item 1A'부터 잘라, 가장 긴 구간(목차가 아닌 본문)을 쓴다."""
    starts = [m.start() for m in ITEM_1A.finditer(text)]
    best = ""
    for end in (m.start() for m in ITEM_1B.finditer(text)):
        before = [s for s in starts if s < end]
        if before and end - before[-1] > len(best):
            best = text[before[-1]:end]
    return best


def sentences(section):
    parts = re.split(r"(?<=[.!?])\s+(?=[A-Z\"“])", re.sub(r"\s+", " ", section))
    return [p.strip() for p in parts if len(p.split()) >= 5]


def words(s):
    return set(re.findall(r"[a-z0-9]+", s.lower()))


def jaccard(a, b):
    return len(a & b) / len(a | b) if a | b else 0.0


def new_sentences(old_text, new_text):
    old = [words(s) for s in sentences(risk_section(old_text))]
    cur = sentences(risk_section(new_text))
    found = []
    for s in cur:
        w = words(s)
        best = max((jaccard(w, o) for o in old), default=0.0)
        if best < THRESHOLD:
            found.append((best, s))
    return len(old), len(cur), found


def main():
    pick = {a.upper() for a in sys.argv[1:]}
    files = defaultdict(list)
    for f in sorted((DATA / "10k").glob("*_*.txt")):
        files[f.stem.split("_")[0]].append(f)
    for ticker, (old_f, new_f) in sorted((t, sorted(fs)[-2:]) for t, fs in files.items() if len(fs) >= 2):
        if pick and ticker not in pick:
            continue
        n_old, n_new, found = new_sentences(old_f.read_text(encoding="utf-8"), new_f.read_text(encoding="utf-8"))
        print(f"== {ticker} {old_f.stem.split('_')[1]} -> {new_f.stem.split('_')[1]}: "
              f"작년 {n_old}문장, 올해 {n_new}문장, 새 문장 {len(found)}개")
        for score, s in found[:5]:
            print(f"  [{score:.2f}] {s[:160]}")


if __name__ == "__main__":
    main()
