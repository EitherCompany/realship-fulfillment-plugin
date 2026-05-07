#!/usr/bin/env python3
"""
v0.3.1 반자동화 풀필먼트 매핑 파이프라인

사용자 confirm 게이트 강제:
- --report-only: 검증 보고서만 출력, 풀필먼트 엑셀 생성 X
- 자동 stop 조건 트리거 시 ABORT
- fixture 테스트 통과 의무
"""
import openpyxl, json, datetime, re, warnings, argparse, sys
from collections import Counter, defaultdict
from openpyxl import Workbook

warnings.filterwarnings('ignore')

HEADERS = ['상품고유코드','판매상품명','수량','배송방식','주문자 이름','받는분 이름','전화번호1','전화번호2',
           '우편번호','주소1','주소2','배송메세지','주문번호','관리메모1','관리메모2','관리메모3','관리메모4','관리메모5',
           '상품별 메모1','상품별 메모2','상품별 메모3','발주 타입','출고희망일']

SR = re.compile(r'(\d{3})\s*-\s*(\d{3})')
LUMI = {f'E004003{n:02d}' for n in range(43, 55)}

NUTRI = [('상어연골',None,'N00200000'),('글루타치온',None,'N00200002'),('브로멜라인',None,'N00200003'),
         ('멜라토닌','5mg','N00200004'),('멜라토닌','2mg','N00200001'),('멜라토닌',None,'N00200004'),
         ('알파CD',None,'N00200005'),('알파시클로덱스트린',None,'N00200005'),('알파씨디',None,'N00200005'),('알파시디',None,'N00200005'),
         ('콜린','미오이노시톨','N00200006'),('비오틴',None,'N00200007'),
         ('초임계 알티지',None,'N00200008'),('rTG',None,'N00200008'),('오메가3',None,'N00200008'),
         ('PS70',None,'N00200009'),('포스파티딜세린',None,'N00200009'),
         ('밀크씨슬',None,'N00200012'),('비타민D',None,'N00200013'),('멀티비타민',None,'N00200014'),
         ('루테인',None,'N00200015'),('마그네슘',None,'N00200016'),('비타민B',None,'N00200017')]


def build_sku_matrix(sku_path):
    """SKU 마스터 → (카테고리, brand, 색상, 사이즈) 매트릭스. brand 분리."""
    wb = openpyxl.load_workbook(sku_path, data_only=True); ws = wb.active
    sku, mat = {}, {}
    for r in range(2, ws.max_row + 1):
        code = ws.cell(r, 3).value; name = ws.cell(r, 4).value
        if not (code and name and (str(code).startswith('E') or str(code).startswith('N'))):
            continue
        sku[code] = name
        if '구름' in name and '깔창' in name:
            m = SR.search(name)
            if m:
                size = f'{m.group(1)}-{m.group(2)}'
                color = '블랙' if '블랙' in name else '그레이' if '그레이' in name else None
                if color:
                    brand = '루미솔' if '루미솔' in name else '이더'
                    mat[('구름깔창', brand, color, size)] = code
        if ('벌집' in name or '메쉬' in name) and '깔창' in name:
            m = SR.search(name)
            if m:
                mat[('벌집깔창', None, None, f'{m.group(1)}-{m.group(2)}')] = code
        if '양말' in name and '클린인테크' in name:
            for c in ['화이트','그레이','블랙']:
                if f'{c} 20p' in name:
                    mat[('양말', None, c, None)] = code
        if '글램루아' in name:
            sm = re.search(r',\s*(L|M|S|XL)\s*$', name)
            if sm:
                seg = name.split(',')
                if len(seg) >= 3:
                    mat[('글램루아', None, seg[-2].strip(), sm.group(1))] = code
        if '아쿠아슈즈' in name:
            sm = SR.search(name)
            if sm:
                size = f'{sm.group(1)}-{sm.group(2)}'
                for c in ['블랙','블루','퍼플','피치','네이비','그레이']:
                    if c in name:
                        mat[('아쿠아슈즈', None, c, size)] = code; break
    return sku, mat


def map_to_code(prod_name, option, mat):
    """매핑 (brand 분리, 베개 순서, 영양제 키워드)"""
    if not prod_name: return None, None
    text = f'{prod_name} {option or ""}'
    # 1) 구름깔창 (이더/루미솔 brand 분리)
    if '구름' in text and ('깔창' in text or '쿠션' in text):
        m = SR.search(text)
        if m:
            size = f'{m.group(1)}-{m.group(2)}'
            for color in ['블랙', '그레이']:
                if (color in text or
                    (color=='블랙' and ('검은색' in text or '검정' in text)) or
                    (color=='그레이' and '회색' in text)):
                    brand = '루미솔' if '루미솔' in text else '이더'
                    k = ('구름깔창', brand, color, size)
                    if k in mat:
                        return mat[k], f'구름깔창 {brand} {color} {size}'
    # 2) 벌집/메쉬
    if ('벌집' in text or '메쉬' in text or '평발' in text) and '깔창' in text:
        m = SR.search(text)
        if m:
            size = f'{m.group(1)}-{m.group(2)}'
            k = ('벌집깔창', None, None, size)
            if k in mat: return mat[k], f'벌집깔창 {size}'
    # 3) 양말
    if '양말' in text:
        for c in ['화이트', '블랙', '그레이']:
            if c in text:
                k = ('양말', None, c, None)
                if k in mat: return mat[k], f'양말 {c}'
    # 4) 글램루아
    if '글램루아' in text:
        sm = re.search(r'사이즈[: ]*([SMLXL]+)', option or '')
        if not sm: sm = re.search(r'\b(L|M|S|XL)\b', option or '')
        size = sm.group(1) if sm else None
        opt2 = (option or '').replace(' ', '')
        if size:
            opt_tok = set(re.findall(r'(블랙|스킨|그레이|화이트)', opt2))
            for k, v in mat.items():
                if k[0] != '글램루아' or k[3] != size: continue
                code_tok = set(re.findall(r'(블랙|스킨|그레이|화이트)', k[2]))
                if opt_tok == code_tok:
                    return v, f'글램루아 {k[2]} {size}'
    # 5) 아쿠아슈즈
    if '아쿠아슈즈' in text:
        m = SR.search(text)
        if m:
            size = f'{m.group(1)}-{m.group(2)}'
            for c in ['블랙','블루','퍼플','피치','네이비','그레이']:
                if c in text:
                    k = ('아쿠아슈즈', None, c, size)
                    if k in mat: return mat[k], f'아쿠아슈즈 {c} {size}'
    # 6) 베개 (덴코 → 슬루나 → 경추)
    if '베개' in text:
        if '덴코' in text: return 'E00400497', '덴코 호텔베개'
        if '슬루나' in text and '호텔' in text: return 'E00400013', '슬루나 호텔베개'
        if '경추' in text: return 'E00400015', '이더 경추베개'
    # 7) 가드웰
    if '가드웰' in text and '무릎' in text and '양쪽' in text:
        return 'E00400484', '가드웰 무릎 양쪽'
    # 8) 영양제
    for kw1, kw2, code in NUTRI:
        if kw1 in text and (kw2 is None or kw2 in text):
            return code, f'뉴트리 {kw1}'
    return None, None


def setmult(option):
    """수량 룰 v3.1 — SKU가 N+N 셋트인 깔창류 반영."""
    if not option: return 1
    s = str(option)
    # N+N (SKU가 1+1 셋트라 ×N)
    m = re.search(r'(\d+)\s*\+\s*(\d+)', s)
    if m and m.group(1) == m.group(2):
        return int(m.group(1))
    # 영양제 N개
    m = re.search(r'할인이벤트[:\s]*(\d+)개', s)
    if m: return int(m.group(1))
    m = re.search(r'(\d+)개\s*:', s)
    if m: return int(m.group(1))
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
    if mall == '쿠팡' and '%' in addr: return 'binbox'
    if mall == '스마트스토어' and '문 앞에 놓아주세요!' in msg: return 'binbox'
    return 'realship'


def add_dong_ho_comma(addr):
    if not addr: return addr
    addr = str(addr).strip()
    pat = re.compile(r'(\s+)(\d+동\s*\d*호?\b|[A-Z]+[\d\-]*동\b|\d+호\b)')
    m = pat.search(addr)
    if not m: return addr
    idx = m.start(1)
    if idx == 0 or addr[idx-1] == ',': return addr
    return addr[:idx] + ',' + addr[idx:]


def norm_zip(z):
    s = str(z or '').strip()
    return s if s and s != 'None' else '00000'


def norm_phone(p):
    s = str(p or '').strip()
    return s if s and s != 'None' else '010-0000-0000'


def norm_addr(a):
    s = str(a or '').strip()
    return add_dong_ho_comma(s) if s else '주소 사방넷 자동입력'


def to_xlsx(rows, sku, output_path):
    wb = Workbook(); ws = wb.active
    ws.append(HEADERS)
    for r in rows:
        ws.append([
            r['_code'],
            sku.get(r['_code'], r.get('_match_name', '')),
            r.get('_qty', r.get('수량') or 1),
            '택배', '',
            (r.get('수취인명') or '').strip(),
            norm_phone(r.get('수취인전화번호1')), '',
            norm_zip(r.get('수취인우편번호(1)')),
            norm_addr(r.get('수취인주소(4)')), '',
            (r.get('배송메세지') or '').strip(),
            str(r.get('주문번호(쇼핑몰)') or '').strip(),
            '','','','','','','','','', ''  # 출고희망일 빈칸
        ])
    wb.save(output_path)


def run_fixture_tests(sku, mat):
    """매 사이클 시작 시 fixture 통과 의무. 실패 시 ABORT."""
    fixtures = [
        # (상품명, 옵션, 예상 코드, 예상 수량(ordQty=1 가정))
        ('바디인솔 1+1 기능성 구름 푹신한 쫀쫀 쿠션 신발 깔창', '색상: 그레이 / 사이즈: 245-250 / 세트: 1+1(2세트 기본할인)', 'E00400008', 1),
        ('바디인솔 1+1 기능성 구름 푹신한 쫀쫀 쿠션 신발 깔창', '색상: 블랙 / 사이즈: 235-240 / 세트: 2+2(4세트 15%)', 'E00400003', 2),
        ('덴코 일자목 거북목 경추베개 호텔베개', '할인이벤트: 2개 : 10% 추가할인', 'E00400497', 2),
        ('이더커머스 목편한 클라우드 경추 베개', '할인이벤트: 1개 : 기본할인', 'E00400015', 1),
        ('가드웰 무릎보호대 마사지볼 세트', '양쪽', 'E00400484', 1),
        ('뉴트리정 글루타치온 57000', '할인이벤트: 3개 : 15% 추가할인', 'N00200002', 3),
        ('뉴트리정 비오틴', '할인이벤트: 1개 : 기본할인', 'N00200007', 1),
        ('뉴트리정 PS70 포스파티딜세린', '할인이벤트: 6개 : 30% 추가할인', 'N00200009', 6),
    ]
    failures = []
    for prod, opt, exp_code, exp_qty in fixtures:
        code, _ = map_to_code(prod, opt, mat)
        qty = 1 * setmult(opt)
        if code != exp_code or qty != exp_qty:
            failures.append((prod[:40], opt[:40], f'code={code}/{exp_code}', f'qty={qty}/{exp_qty}'))
    return failures


def build_report(rows, pre_cycle, binbox, realship, mapped, unmapped, state_skip, final, ether, nutri, dup, sku):
    code_dist = Counter(r['_code'] for r in final)
    total = len(final)
    
    # 자동 stop 조건 검사
    stops = []
    if len(unmapped) > 0:
        stops.append(f'미매핑 {len(unmapped)}건')
    if total > 0:
        max_pct = code_dist.most_common(1)[0][1] / total * 100 if code_dist else 0
        if max_pct > 50 and code_dist.most_common(1)[0][0] != 'E00400023':  # 양말 화이트는 베스트셀러 화이트리스트
            stops.append(f'한 코드 {max_pct:.1f}% 몰림')
    lumi_count = sum(v for k, v in code_dist.items() if k in LUMI)
    if lumi_count > 0:
        stops.append(f'루미솔 코드 {lumi_count}건 — 정상 케이스인지 확인')
    if len(dup) > 0:
        stops.append(f'cross-validate 중복 {len(dup)}건')
    
    # 우편번호 fallback 비율
    zip_fb = sum(1 for r in final if norm_zip(r.get('수취인우편번호(1)')) == '00000')
    zip_pct = zip_fb / total * 100 if total else 0
    if zip_pct > 5:
        stops.append(f'우편번호 00000 fallback {zip_pct:.1f}%')
    
    # 수량 매트릭스
    qty_matrix_NN = defaultdict(int)
    qty_matrix_N = defaultdict(int)
    for r in final:
        opt = r.get('옵션(수집)') or ''
        m = re.search(r'(\d+)\s*\+\s*(\d+)', opt)
        if m and m.group(1) == m.group(2):
            n = int(m.group(1))
            ord_q = r.get('수량') or 1
            qty_matrix_NN[(f'{n}+{n}', float(ord_q), float(r['_qty']))] += 1
        m = re.search(r'(\d+)개\s*:', opt)
        if m:
            n = int(m.group(1))
            ord_q = r.get('수량') or 1
            qty_matrix_N[(f'{n}개', float(ord_q), float(r['_qty']))] += 1
    
    return {
        'totals': {
            'all': len(rows),
            'pre_cycle_skip': len(pre_cycle),
            'binbox_skip': len(binbox),
            'realship': len(realship),
            'mapped': len(mapped),
            'unmapped': len(unmapped),
            'state_skip': len(state_skip),
            'crossvalidate_dup': len(dup),
            'final': len(final),
            'ether': len(ether),
            'nutri': len(nutri),
        },
        'top_codes': code_dist.most_common(10),
        'lumisol_count': lumi_count,
        'zip_fallback_pct': zip_pct,
        'qty_matrix_NN': {f'{k[0]}, ordQty={k[1]}, 풀필={k[2]}': v for k,v in qty_matrix_NN.items()},
        'qty_matrix_N': {f'{k[0]}, ordQty={k[1]}, 풀필={k[2]}': v for k,v in qty_matrix_N.items()},
        'unmapped_samples': [
            {'product': r.get('상품명(수집)'), 'option': r.get('옵션(수집)')}
            for r in unmapped[:10]
        ],
        'auto_stop_reasons': stops,
        'should_proceed': len(stops) == 0,
    }


def main():
    p = argparse.ArgumentParser()
    p.add_argument('--orders', required=True)
    p.add_argument('--sku-master')
    p.add_argument('--fulfillment-history', nargs='*', help='풀필먼트 발주조회 엑셀(들) — cross-validate용')
    p.add_argument('--state', required=True)
    p.add_argument('--cycle-start', required=True)
    p.add_argument('--report-only', action='store_true', help='검증 보고서만 출력. 풀필먼트 엑셀 생성 X')
    p.add_argument('--output-ether')
    p.add_argument('--output-nutri')
    p.add_argument('--output-report')
    args = p.parse_args()

    # SKU
    sku, mat = build_sku_matrix(args.sku_master) if args.sku_master else ({}, {})
    
    # Fixture 테스트 — 통과 못하면 ABORT
    if mat:
        failures = run_fixture_tests(sku, mat)
        if failures:
            print('🚨 FIXTURE 테스트 실패. ABORT.')
            for f in failures: print(f'  {f}')
            sys.exit(1)
        print('✅ Fixture 테스트 통과')

    # 엑셀 로드
    wb = openpyxl.load_workbook(args.orders, data_only=True); ws = wb.active
    h = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    rows = [{h[c]: ws.cell(r, c+1).value for c in range(len(h)) if h[c]}
            for r in range(2, ws.max_row + 1)]

    cycle_start = datetime.datetime.strptime(args.cycle_start, '%Y-%m-%d %H:%M')

    # state
    try:
        with open(args.state, encoding='utf-8') as f: state = json.load(f)
    except: state = {}
    prev_pairs = set()
    for k, v in state.items():
        if isinstance(v, dict):
            for plist in [v.get('shma_code_pairs', []), v.get('ether_code_pairs', []), v.get('nutri_code_pairs', [])]:
                for p_ in plist:
                    if isinstance(p_, list) and len(p_) == 2:
                        prev_pairs.add(tuple(p_))

    # 분류
    pre_cycle, binbox, realship = [], [], []
    for r in rows:
        if is_pre_cycle(r, cycle_start): pre_cycle.append(r); continue
        c = classify(r)
        (binbox if c == 'binbox' else realship).append(r)

    # 매핑
    mapped, unmapped = [], []
    for r in realship:
        code, name = map_to_code(r.get('상품명(수집)'), r.get('옵션(수집)'), mat)
        if code:
            r['_code'] = code; r['_match_name'] = name
            r['_qty'] = (r.get('수량') or 1) * setmult(r.get('옵션(수집)'))
            mapped.append(r)
        else:
            unmapped.append(r)

    # state 차집합 (사용자 수동 처리 사이클은 거짓 가능성 — cross-validate 우선)
    state_skip, after_state = [], []
    for r in mapped:
        pair = (str(r.get('주문번호(쇼핑몰)','')), r['_code'])
        if pair in prev_pairs: state_skip.append(r)
        else: after_state.append(r)

    # cross-validate
    fulfill = []
    if args.fulfillment_history:
        try:
            import msoffcrypto, io
            for path in args.fulfillment_history:
                with open(path, 'rb') as f:
                    of = msoffcrypto.OfficeFile(f); of.load_key(password='dlejrhddyd1!')
                    buf = io.BytesIO(); of.decrypt(buf); buf.seek(0)
                fws = openpyxl.load_workbook(buf, data_only=True).active
                fh = [fws.cell(1, c).value for c in range(1, fws.max_column+1)]
                for rr in range(2, fws.max_row+1):
                    fulfill.append({fh[c]: fws.cell(rr, c+1).value for c in range(len(fh)) if fh[c]})
        except Exception as e:
            print(f'⚠️ 발주조회 로드 실패: {e}')

    def npn(p): return str(p or '').strip().replace('-','').replace(' ','')
    fA, fB = defaultdict(list), defaultdict(list)
    for r in fulfill:
        on = str(r.get('주문번호') or '').strip()
        cd = str(r.get('고유코드') or '').strip()
        nm = str(r.get('받는분 이름') or '').strip()
        ph = npn(r.get('전화번호1'))
        if on and cd: fA[(on, cd)].append(r)
        if nm and ph and cd: fB[(nm, ph, cd)].append(r)

    dup, final = [], []
    for r in after_state:
        kA = (str(r.get('주문번호(쇼핑몰)','')).strip(), r['_code'])
        kB = ((r.get('수취인명') or '').strip(), npn(r.get('수취인전화번호1')), r['_code'])
        if (fA and kA in fA) or (fB and kB in fB): dup.append(r)
        else: final.append(r)

    ether = [r for r in final if r['_code'].startswith('E')]
    nutri = [r for r in final if r['_code'].startswith('N')]

    # 검증 보고서
    report = build_report(rows, pre_cycle, binbox, realship, mapped, unmapped, state_skip, final, ether, nutri, dup, sku)
    if args.output_report:
        with open(args.output_report, 'w', encoding='utf-8') as f:
            json.dump(report, f, ensure_ascii=False, indent=2, default=str)

    # 콘솔 출력
    print(f'\n[검증 보고서]')
    print(f'  전체 {report["totals"]["all"]} | 사이클이전 {report["totals"]["pre_cycle_skip"]} | 빈박스 {report["totals"]["binbox_skip"]} | 실배송 {report["totals"]["realship"]}')
    print(f'  매핑 {report["totals"]["mapped"]} / 미매핑 {report["totals"]["unmapped"]}')
    print(f'  state SKIP {report["totals"]["state_skip"]} / cross-validate 중복 {report["totals"]["crossvalidate_dup"]}')
    print(f'  최종 {report["totals"]["final"]} (이더 {report["totals"]["ether"]} / 뉴트리 {report["totals"]["nutri"]})')
    print(f'  루미솔 카운트: {report["lumisol_count"]} | 우편 fallback: {report["zip_fallback_pct"]:.1f}%')
    print(f'  코드 TOP 5: {report["top_codes"][:5]}')
    print(f'\n자동 stop 조건: {len(report["auto_stop_reasons"])}건')
    for s in report["auto_stop_reasons"]: print(f'  ⚠️ {s}')

    # report-only 모드 → 종료
    if args.report_only:
        print('\n📋 --report-only 모드 — 풀필먼트 엑셀 생성 안 함. 사용자 confirm 후 다시 실행.')
        return

    # 자동 stop
    if not report['should_proceed']:
        print('\n🚨 자동 stop 조건 발생. 사용자 confirm 받기 전엔 풀필먼트 엑셀 생성 X.')
        sys.exit(2)

    # 풀필먼트 엑셀 생성
    if args.output_ether: to_xlsx(ether, sku, args.output_ether)
    if args.output_nutri: to_xlsx(nutri, sku, args.output_nutri)

    # state 갱신
    today_key = datetime.date.today().isoformat()
    state['last_cycle_date'] = today_key
    state[today_key] = {
        'ether_code_pairs': [[str(r.get('주문번호(쇼핑몰)','')), r['_code']] for r in ether],
        'nutri_code_pairs': [[str(r.get('주문번호(쇼핑몰)','')), r['_code']] for r in nutri],
        'shma_code_pairs': [[str(r.get('주문번호(쇼핑몰)','')), r['_code']] for r in final],
        'binbox_skip_count': len(binbox),
        'pre_cycle_skip_count': len(pre_cycle),
        'unmapped_count': len(unmapped),
        'crossvalidate_dup_count': len(dup),
        'cycle_start': args.cycle_start,
        'complete_time': datetime.datetime.now().strftime('%Y-%m-%d %H:%M KST'),
    }
    with open(args.state, 'w', encoding='utf-8') as f:
        json.dump(state, f, ensure_ascii=False, indent=2)


if __name__ == '__main__':
    main()
