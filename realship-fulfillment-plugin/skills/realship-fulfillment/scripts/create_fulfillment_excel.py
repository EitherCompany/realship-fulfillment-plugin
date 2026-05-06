#!/usr/bin/env python3
"""
v0.3.0 반자동화 풀필먼트 매핑 파이프라인

사용자가 사방넷에서 다운받은 주문서확인처리 엑셀을 입력받아:
1. 사이클 이전 SKIP
2. 빈박스 SKIP (쿠팡 % / 스마트 '문 앞에 놓아주세요!')
3. 실배송 추출 + 상품매핑 (SKU 마스터 매트릭스 + 영양제 키워드)
4. state 차집합 가드
5. 이더(E*) / 뉴트리(N*) 풀필먼트 업로드 엑셀 생성
6. 특이사항 리포트 (JSON)
"""
import openpyxl, json, datetime, re, warnings, argparse
from collections import Counter, defaultdict
from openpyxl import Workbook

warnings.filterwarnings('ignore')

# 풀필먼트 발주 엑셀 헤더 (23 컬럼)
HEADERS = [
    '상품고유코드', '판매상품명', '수량', '배송방식',
    '주문자 이름', '받는분 이름', '전화번호1', '전화번호2',
    '우편번호', '주소1', '주소2', '배송메세지', '주문번호',
    '관리메모1', '관리메모2', '관리메모3', '관리메모4', '관리메모5',
    '상품별 메모1', '상품별 메모2', '상품별 메모3',
    '발주 타입', '출고희망일'
]

SIZE_RE = re.compile(r'(\d{3})\s*-\s*(\d{3})')
SET_PLUS_RE = re.compile(r'(\d+)\s*\+\s*\1')
SET_NUM_RE = re.compile(r'(\d+)세트')


def build_sku_matrix(sku_path):
    """풀필먼트 SKU 마스터(재고조회) → 카테고리별 매트릭스"""
    wb = openpyxl.load_workbook(sku_path, data_only=True)
    ws = wb.active
    sku, mat = {}, {}
    for r in range(2, ws.max_row + 1):
        code = ws.cell(r, 3).value
        name = ws.cell(r, 4).value
        if not (code and name and (str(code).startswith('E') or str(code).startswith('N'))):
            continue
        sku[code] = name
        # 구름깔창
        if '구름' in name and '깔창' in name:
            m = SIZE_RE.search(name)
            if m:
                size = f'{m.group(1)}-{m.group(2)}'
                color = '블랙' if '블랙' in name else '그레이' if '그레이' in name else None
                if color:
                    mat[('구름깔창', color, size)] = code
        # 벌집/메쉬 깔창
        if ('벌집' in name or '메쉬' in name) and '깔창' in name:
            m = SIZE_RE.search(name)
            if m:
                mat[('벌집깔창', None, f'{m.group(1)}-{m.group(2)}')] = code
        # 양말
        if '양말' in name and '클린인테크' in name:
            for color in ['화이트', '그레이', '블랙']:
                if f'{color} 20p' in name:
                    mat[('양말', color, None)] = code
        # 글램루아
        if '글램루아' in name:
            sm = re.search(r',\s*(L|M|S|XL)\s*$', name)
            if sm:
                seg = name.split(',')
                if len(seg) >= 3:
                    mat[('글램루아', seg[-2].strip(), sm.group(1))] = code
        # 아쿠아슈즈
        if '아쿠아슈즈' in name:
            sm = SIZE_RE.search(name)
            if sm:
                size = f'{sm.group(1)}-{sm.group(2)}'
                for c in ['블랙', '블루', '퍼플', '피치', '네이비', '그레이']:
                    if c in name:
                        mat[('아쿠아슈즈', c, size)] = code
                        break
    return sku, mat


NUTRI_RULES = [
    ('상어연골', None, 'N00200000'),
    ('글루타치온', None, 'N00200002'),
    ('브로멜라인', None, 'N00200003'),
    ('멜라토닌', '5mg', 'N00200004'),
    ('멜라토닌', '2mg', 'N00200001'),
    ('멜라토닌', None, 'N00200004'),
    ('알파CD', None, 'N00200005'),
    ('알파시클로덱스트린', None, 'N00200005'),
    ('알파씨디', None, 'N00200005'),
    ('알파시디', None, 'N00200005'),
    ('콜린', '미오이노시톨', 'N00200006'),
    ('비오틴', None, 'N00200007'),
    ('초임계 알티지', None, 'N00200008'),
    ('rTG', None, 'N00200008'),
    ('오메가3', None, 'N00200008'),
    ('PS70', None, 'N00200009'),
    ('포스파티딜세린', None, 'N00200009'),
    ('밀크씨슬', None, 'N00200012'),
    ('비타민D', None, 'N00200013'),
    ('멀티비타민', None, 'N00200014'),
    ('루테인', None, 'N00200015'),
    ('마그네슘', None, 'N00200016'),
    ('비타민B', None, 'N00200017'),
]


def map_to_code(prod_name, option, mat_index):
    """상품명+옵션 → 풀필먼트 코드. 색상은 원문/정규화 양쪽 검색."""
    if not prod_name:
        return None, None
    text = f'{prod_name} {option or ""}'
    # 1) 구름깔창
    if '구름' in text and ('깔창' in text or '쿠션' in text):
        m = SIZE_RE.search(text)
        if m:
            size = f'{m.group(1)}-{m.group(2)}'
            for color in ['블랙', '그레이']:
                if color in text or (color == '블랙' and ('검은색' in text or '검정' in text)) or (color == '그레이' and '회색' in text):
                    k = ('구름깔창', color, size)
                    if k in mat_index:
                        return mat_index[k], f'구름깔창 {color} {size}'
    # 2) 벌집/메쉬
    if ('벌집' in text or '메쉬' in text or '평발' in text) and '깔창' in text:
        m = SIZE_RE.search(text)
        if m:
            size = f'{m.group(1)}-{m.group(2)}'
            k = ('벌집깔창', None, size)
            if k in mat_index:
                return mat_index[k], f'벌집깔창 {size}'
    # 3) 양말
    if '양말' in text:
        for color in ['화이트', '블랙', '그레이']:
            if color in text:
                k = ('양말', color, None)
                if k in mat_index:
                    return mat_index[k], f'양말 {color} 20p'
    # 4) 글램루아
    if '글램루아' in text:
        sm = re.search(r'사이즈[: ]*([SMLXL]+)', option or '')
        if not sm:
            sm = re.search(r'\b(L|M|S|XL)\b', option or '')
        size = sm.group(1) if sm else None
        opt = (option or '').replace(' ', '')
        if size:
            opt_tokens = set(re.findall(r'(블랙|스킨|그레이|화이트)', opt))
            for k, v in mat_index.items():
                if k[0] != '글램루아' or k[2] != size:
                    continue
                code_tokens = set(re.findall(r'(블랙|스킨|그레이|화이트)', k[1]))
                if opt_tokens == code_tokens:
                    return v, f'글램루아 {k[1]} {size}'
    # 5) 아쿠아슈즈
    if '아쿠아슈즈' in text:
        m = SIZE_RE.search(text)
        if m:
            size = f'{m.group(1)}-{m.group(2)}'
            for color in ['블랙', '블루', '퍼플', '피치', '네이비', '그레이']:
                if color in text:
                    k = ('아쿠아슈즈', color, size)
                    if k in mat_index:
                        return mat_index[k], f'아쿠아슈즈 {color} {size}'
    # 6) 베개
    if '경추' in text and '베개' in text:
        return 'E00400015', '경추베개'
    if '덴코' in text and '베개' in text:
        return 'E00400497', '덴코베개'
    # 7) 가드웰
    if '가드웰' in text and '무릎' in text and '양쪽' in text:
        return 'E00400484', '가드웰 무릎보호대 양쪽'
    # 8) 영양제
    for kw1, kw2, code in NUTRI_RULES:
        if kw1 in text and (kw2 is None or kw2 in text):
            return code, f'뉴트리 {kw1}'
    return None, None


def set_multiplier(option):
    if not option:
        return 1
    o = str(option)
    m = SET_PLUS_RE.search(o)
    if m:
        return int(m.group(1)) * 2
    m = SET_NUM_RE.search(o)
    if m:
        return int(m.group(1))
    return 1


def is_pre_cycle(r, cycle_start):
    od = r.get('주문일시(YYYY-MM-DD HH:MM)')
    if isinstance(od, datetime.datetime):
        return od < cycle_start
    if isinstance(od, str):
        try:
            return datetime.datetime.strptime(od[:16], '%Y-%m-%d %H:%M') < cycle_start
        except Exception:
            return False
    return False


def classify(r):
    mall = r.get('쇼핑몰명(1)') or ''
    addr = r.get('수취인주소(4)') or ''
    msg = r.get('배송메세지') or ''
    if mall == '쿠팡' and '%' in addr:
        return 'binbox'
    if mall == '스마트스토어' and '문 앞에 놓아주세요!' in msg:
        return 'binbox'
    return 'realship'


def normalize_zip(z):
    z = str(z or '').strip()
    if not z or z == 'None' or '-' in z or len(z) > 5:
        return '00000'
    return z


def normalize_phone(p):
    p = str(p or '').strip()
    if not p or p == 'None':
        return '010-0000-0000'
    return p


def normalize_addr(a):
    a = str(a or '').strip()
    if not a:
        return '주소 사방넷 자동입력'
    return a


def to_xlsx(rows, sku, output_path, tomorrow):
    wb = Workbook(); ws = wb.active
    ws.append(HEADERS)
    for r in rows:
        ws.append([
            r['_code'],
            sku.get(r['_code'], r.get('_match_name', '')),
            r.get('_qty', r.get('수량') or 1),
            '택배',
            '',
            (r.get('수취인명') or '').strip(),
            normalize_phone(r.get('수취인전화번호1')),
            '',
            normalize_zip(r.get('수취인우편번호(1)')),
            normalize_addr(r.get('수취인주소(4)')),
            '',
            (r.get('배송메세지') or '').strip(),
            str(r.get('주문번호(쇼핑몰)') or '').strip(),
            '', '', '', '', '',
            '', '', '',
            '', tomorrow
        ])
    wb.save(output_path)


def main():
    p = argparse.ArgumentParser()
    p.add_argument('--orders', required=True)
    p.add_argument('--sku-master')
    p.add_argument('--state', required=True)
    p.add_argument('--cycle-start', required=True)
    p.add_argument('--output-ether', required=True)
    p.add_argument('--output-nutri', required=True)
    p.add_argument('--report')
    args = p.parse_args()

    wb = openpyxl.load_workbook(args.orders, data_only=True)
    ws = wb.active
    h = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    rows = [{h[c]: ws.cell(r, c + 1).value for c in range(len(h)) if h[c]}
            for r in range(2, ws.max_row + 1)]

    sku, mat = build_sku_matrix(args.sku_master) if args.sku_master else ({}, {})
    cycle_start = datetime.datetime.strptime(args.cycle_start, '%Y-%m-%d %H:%M')
    tomorrow = (datetime.date.today() + datetime.timedelta(days=1)).isoformat()

    try:
        with open(args.state, encoding='utf-8') as f:
            state = json.load(f)
    except Exception:
        state = {}
    prev_pairs = set()
    for k, v in state.items():
        if isinstance(v, dict):
            for plist in [v.get('shma_code_pairs', []),
                          v.get('ether_code_pairs', []),
                          v.get('nutri_code_pairs', [])]:
                for p_ in plist:
                    if isinstance(p_, list) and len(p_) == 2:
                        prev_pairs.add(tuple(p_))

    pre_cycle, binbox, realship = [], [], []
    for r in rows:
        if is_pre_cycle(r, cycle_start):
            pre_cycle.append(r); continue
        c = classify(r)
        (binbox if c == 'binbox' else realship).append(r)

    mapped, unmapped = [], []
    for r in realship:
        code, name = map_to_code(r.get('상품명(수집)'), r.get('옵션(수집)'), mat)
        if code:
            r['_code'] = code; r['_match_name'] = name
            r['_qty'] = (r.get('수량') or 1) * set_multiplier(r.get('옵션(수집)'))
            mapped.append(r)
        else:
            unmapped.append(r)

    state_skip, final = [], []
    for r in mapped:
        pair = (str(r.get('주문번호(쇼핑몰)', '')), r['_code'])
        if pair in prev_pairs:
            state_skip.append(r)
        else:
            final.append(r)

    ether = [r for r in final if r['_code'].startswith('E')]
    nutri = [r for r in final if r['_code'].startswith('N')]
    to_xlsx(ether, sku, args.output_ether, tomorrow)
    to_xlsx(nutri, sku, args.output_nutri, tomorrow)

    today_key = datetime.date.today().isoformat()
    state['last_cycle_date'] = today_key
    state[today_key] = {
        'ether_code_pairs': [[str(r.get('주문번호(쇼핑몰)', '')), r['_code']] for r in ether],
        'nutri_code_pairs': [[str(r.get('주문번호(쇼핑몰)', '')), r['_code']] for r in nutri],
        'shma_code_pairs': [[str(r.get('주문번호(쇼핑몰)', '')), r['_code']] for r in final],
        'binbox_skip_count': len(binbox),
        'pre_cycle_skip_count': len(pre_cycle),
        'unmapped_count': len(unmapped),
        'complete_time': datetime.datetime.now().strftime('%Y-%m-%d %H:%M KST'),
    }
    with open(args.state, 'w', encoding='utf-8') as f:
        json.dump(state, f, ensure_ascii=False, indent=2)

    if args.report:
        code_dist = Counter(r['_code'] for r in final)
        report = {
            'cycle_date': today_key,
            'totals': {
                'all': len(rows),
                'pre_cycle_skip': len(pre_cycle),
                'binbox_skip': len(binbox),
                'realship': len(realship),
                'mapped': len(mapped),
                'unmapped': len(unmapped),
                'state_skip': len(state_skip),
                'final': len(final),
                'ether': len(ether),
                'nutri': len(nutri),
            },
            'top_codes': code_dist.most_common(20),
        }
        with open(args.report, 'w', encoding='utf-8') as f:
            json.dump(report, f, ensure_ascii=False, indent=2)

    print(f'전체 {len(rows)} | 사이클이전 {len(pre_cycle)} | 빈박스 {len(binbox)} | 실배송 {len(realship)} | 매핑 {len(mapped)} | 미매핑 {len(unmapped)} | state SKIP {len(state_skip)} | 최종 {len(final)} (이더 {len(ether)} / 뉴트리 {len(nutri)})')


if __name__ == '__main__':
    main()
