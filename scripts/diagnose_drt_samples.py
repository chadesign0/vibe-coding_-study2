"""drt 메타 fallback 도입 전 표본 검증.
다양한 유형 키워드 5개에 대해 통합검색 페이지의 drt 메타 추출 →
순서를 그대로 출력. 사용자가 브라우저 화면과 비교해 일치 여부 확인.

NAVER throttle 대응: 매 요청마다 새 session + 30초 쿨다운."""
import sys, io, re, time, random
from urllib.parse import quote_plus
from pathlib import Path
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.path.insert(0, str(Path(__file__).parent))

import requests
from build_month import decode_response_text, SEARCH_HEADERS

# 다양한 유형 표본 — 클래스병원 keyword 풀에서 선별
SAMPLES = [
    ('자기동네 광역',   '안산정형외과'),
    ('자기동네 핵심',   '고잔정형외과'),
    ('역세권 일반',     '중앙역정형외과'),
    ('인접동',          '본오동정형외과'),
    ('주변 동',         '성포동정형외과'),
]

def fetch_fresh(query: str):
    """매 요청마다 새 session — throttle 회피."""
    sess = requests.Session()
    sess.headers.update(SEARCH_HEADERS)
    url = 'https://search.naver.com/search.naver?where=nexearch&sm=tab_jum&ssc=tab.nx.all&query=' + quote_plus(query)
    try:
        r = sess.get(url, timeout=30)
        if r.status_code != 200:
            return None, r.status_code
        return decode_response_text(r), 200
    except Exception as e:
        return None, str(e)

def extract_drt(html: str):
    return re.findall(r'"id":"(\d+)","dbType":"drt","name":"([^"]+)"', html or '')

print(f"\n{'='*78}")
print(f"  drt 메타 표본 검증 — 10개 키워드 통합검색 페이지의 plat 카드 순서")
print(f"  사용자가 브라우저에서 동일 키워드 검색 후 첫 화면의 플레이스 카드 순서와 비교")
print(f"{'='*78}\n")

for tag, kw in SAMPLES:
    ht, status = fetch_fresh(kw)
    if not ht:
        print(f"[{tag:12s}] {kw}  → fetch 실패 (status={status})\n")
        time.sleep(30)
        continue
    metas = extract_drt(ht)
    print(f"[{tag:12s}] {kw}  ({len(metas)}건)")
    for i, (pid, name) in enumerate(metas[:8], 1):
        is_class = '★' if '클래스병원' in name else ' '
        print(f"    {is_class}{i}: {name}  (id={pid})")
    if not metas:
        print(f"     (drt 메타 0건 — 이 키워드는 plat 카드 미노출 가능성)")
    print()
    time.sleep(30)

print(f"{'='*78}")
print("  검증 방법: 브라우저로 https://search.naver.com/search.naver?query=<키워드>")
print("  접속 후 플레이스 카드 1~8위가 위 출력과 동일 순서인지 확인")
print(f"{'='*78}")
