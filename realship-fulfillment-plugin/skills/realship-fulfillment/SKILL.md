---
name: realship-fulfillment
description: >
  실배송 풀필먼트 자동 등록 스킬 (v0.2.6 — 수량 룰 v3 / cnt_multiplier 채널 분기 / 시리즈 B 룰).
  매일 오후 2시 주문 마감 후 사방넷 → 사방넷 풀필먼트(이더+뉴트리)에 실배송만 자동 등록.
  반드시 이 스킬을 사용해야 하는 경우: "실배송 풀필먼트", "풀필먼트 등록", "실배송 등록",
  "오늘 실배송", "풀필먼트 발주", "실배송 처리해줘", "주문 마감", "2시 마감",
  "풀필먼트에 넣어줘", "실배송 올려줘", "발주등록해줘", "실배송 1차", "실배송 2차",
  "사방넷 발주", "스마트스토어 발주", "발주 엑셀 만들어줘"
---

# 실배송 풀필먼트 자동 등록 (v0.2.6)

## 시스템 정보

### 사방넷 관리자
| 항목 | 값 |
|---|---|
| URL | `https://sbadmin03.sabangnet.co.kr` |
| 로그인 | `eithercompany` / `dlejzja7801!` |
| svcAcntId | `mw159514` |

### 사방넷 풀필먼트 (계정 2개, 비밀번호 통일)
| 계정 | 회사코드 | 아이디 | 비밀번호 | 취급 |
|---|---|---|---|---|
| 이더컴퍼니 (공산품) | `w7298` | `eithercompany` | **`dlejrhddyd1@`** | E-코드 |
| 뉴트리정 (영양제) | `w7298` | `nutrijung` | **`dlejrhddyd1@`** | N-코드 |

URL: `https://wms02.sbfulfillment.co.kr`

⚠️ **v0.2.6 비밀번호 정정**: 이전 SKILL.md에 `dlejrhddyd1!` 로 잘못 기록됐던 것이 사실은 `dlejrhddyd1@` 임이 확인됨 (2026-04-30).

⚠️ **풀필먼트 다운로드 엑셀은 비밀번호 보호** — 위 비번으로 복호화 (msoffcrypto-tool)

---

## v0.2.6 절대 룰 (사고 기반 영구 반영)

1. **shmaOrdNo 단일 키 매칭** — 받는분 이름 매칭 절대 금지 (김채은 cross-account 사고 차단)
2. **🆕 수량 룰 v3** (v0.2.5 v2 → v3 갱신)

   ```
   풀필 수량 = ordQt × set_multiplier(skuNm) × cnt_multiplier(skuNm, 채널)

   set_multiplier(skuNm):
     • skuNm에 'N+N' 패턴 (1+1, 2+2, 3+3) → N
       (1+1 → 1, 2+2 → 2, 3+3 → 3)
     • product_mapping의 세트_배수 dict 기반 (단일 source of truth)
     • 룰에 skip_multiplier:true 있으면 → 1 강제
     • 매칭 없음 → 1

   cnt_multiplier(skuNm, shma_login_id):
     • 쿠팡 (CP) — shma_login_id ∈ {nutrijung, mineflow, cleanintech, edencorporation1}
       'N개입 / N박스 / N통 / N병 / N정 / N개' → N
       예: '[CP] 160g 4박스' → 4
     • 스마트스토어 (SS) — shma_login_id 가 'ncp_' 로 시작
       '할인이벤트:' 키워드 컨텍스트 안의 'N개' → N
       (할인이벤트 키워드 없으면 N개 매칭 금지 — 사이즈/색상의 N과 혼동 방지)
     • 매칭 없음 → 1

   set_multiplier와 cnt_multiplier는 독립. 둘 다 매칭되면 곱하기.
   ```

3. **빈박스 룰** — `'문 앞에 놓아주세요!' in shpmtMsg` 느낌표 정확 매칭만 빈박스. 부분 매칭 금지!
   - 04-29 사고: 부분 매칭으로 9건(스마트스토어 일반 안전 메시지) 누락 직전
4. **사전 일괄확정·001→002 자동 의무화** (Vue VM 트릭)
5. **🆕 풀필먼트 엑셀 생성 후 사용자 confirm 의무 (수량 샘플 검증 추가)**
   1. 사업자별 분리 결과 (이더 E* / 뉴트리 N* / 미분류)
   2. 복수구매 그룹 (같은 받는분+주소) → 합배송 OK 여부
   3. 동성동명 의심 (같은 이름, 다른 주소) → 별도 발송 OK 여부
   4. 미분류 → product_mapping.json 룰 추가
   5. **🆕 1+1 옵션 row 샘플 → 풀필먼트 수량이 1인지**
   6. **🆕 2+2 옵션 row 샘플 → 풀필먼트 수량이 2인지**
   7. **🆕 N박스/N개입 row 샘플 (쿠팡) → 수량이 N인지**
   8. **🆕 ffQty 분포 자동 출력 (`print_ffqty_distribution()`) — 모든 row가 짝수면 1+1 ×2 사고 의심 경고**

6. **04-28 같은 사이클 시작일 이전 주문 자동 SKIP**
7. **state.json: 오늘 풀필먼트 등록 ord_no 저장**
8. **풀필먼트 발주 엑셀 헤더는 반드시 스크립트의 HEADERS 상수 import 사용**
9. **빈 전화/우편/주소 fallback** — 010-0000-0000 / 00000 / "주소 사방넷 자동입력"
10. **출고희망일 자동 = 다음날 (YYYY-MM-DD)**
11. **스마트스토어 일주일치 vs 풀필먼트 엑셀 크로스체킹 의무**
12. **list API 2종 join 패턴 — 엑셀 다운로드 완전 폐지** (v0.2.5)
13. **풀필먼트 등록 직전 ordStsCd 재확인 가드 의무** (v0.2.5)
14. **🆕 일괄확정 / 001→002 변환 후 0건 검증 의무** (v0.2.6 — 04-30 catch)

    각 단계 처리 직후 같은 검색 조건으로 재검색하여 **자료수 0건** 확인.
    0건이 아니면 단계 재시도. 좌표 클릭 누락이 일괄확정 단계에서 1회 발생한 사고 catch 후 추가.

15. **🆕 Python ↔ JS 매핑 함수 단일 source of truth** (v0.2.6 — 37건 catch)

    페이지 JS에서 임시로 만든 매핑 함수가 Python `create_fulfillment_excel.py` 명세와 다르게 구현된 것이 37건 set_multiplier 사고 원인. v0.2.6부터:
    - **Python `get_set_multiplier` / `get_cnt_multiplier` 가 단일 명세**
    - 페이지 JS는 같은 정규식 + 같은 분기 룰 사용 (이 SKILL.md의 명세 그대로)
    - JSON `세트_배수` dict는 Python·JS 모두 source

16. **🆕 SheetJS CDN — jsdelivr 우선** (v0.2.6 — cdnjs partial load 사고 catch)

    cdnjs.cloudflare.com 의 SheetJS는 partial load (window.XLSX는 set되지만 XLSX.utils 가 undefined) 발생.
    `https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js` 우선 사용.

17. **🆕 list API body는 페이지 vm sbForm deep copy 후 일부 필드만 변경** (v0.2.6)

    인라인으로 만든 body는 `code: 99999` 에러 발생. 페이지 vm의 sbForm을 `JSON.parse(JSON.stringify(pv.sbForm))` 으로 deep copy 후 startDate/endDate/pageSize/orderStatus/wyblNoregs 등만 set.

---

## v0.2.6 변경 요약 (2026-04-30 사이클 catch)

| 사고 | 영향 row | 정정 |
|---|---|---|
| 1+1 옵션을 ×2로 곱셈 (페이지 JS) | 37건 | set_multiplier dict 명세 → Python·JS 동일 |
| `[CP] 160g 4박스` → ×1 (Python·JS 모두 박스 단위 미지원) | 1건 | cnt_multiplier 신규 함수 + 채널 분기 |
| 비밀번호 ! → @ | — | 영구 정정 |
| prdNo 100006 시리즈 B (남녀공용 논슬립 워터슈즈, E00400656~) 누락 | 1건 (블루 250-255) | 키워드 룰 1개 추가 (사용자 확정), 나머지는 풀필먼트 마스터 dump 후 보완 |
| 일괄확정 좌표 1회 누락 | 0건 (catch됨) | 단계마다 0건 검증 의무 (룰 14) |
| cdnjs SheetJS partial load | — | jsdelivr 우선 (룰 16) |
| list API body code 99999 | — | sbForm deep copy 룰 17 |

---

## v0.2.5 핵심 우회 (그대로 유지)

엑셀 다운로드 클릭 0번. 두 list API ordNo join.

| API | 응답 핵심 필드 | 누락 필드 |
|---|---|---|
| `/prod-api/customer/order/OrderConfirm/searchOrders` | `shpmtMsg`, `shpmtEtcFldVl`, `shmaNm`, `ordClctFldVl1~4` | `ecptRmteTotAddr`, `rmteZipcd` |
| `/prod-api/customer/order/WaybillInputSku/getWaybillInputSkuLists` | `ecptRmteTotAddr`, `rmteZipcd`, `ecptRmteNm`, `ecptRmteTelNo`, `ecptRmteHndpnNo`, `clctPrdNm`, `clctSkuNm`, `ordQt`, `ordStsCd` | `shpmtMsg` |

`ordNo` join → 100% 완전 데이터셋.

---

## 매 사이클 표준 6단계

### Step 1 — 사방넷 주문수집 (자동 ✅)
`#/order/order-collect-auto` — 7개 mall 일괄 선택 → 주문수집 → 확인.

### Step 2 — 사방넷 일괄확정 (자동 ✅ + 0건 검증 의무)
`#/order/order-decide` — 검색 → 일괄주문확정 → 확인 → 모달 일괄주문확정 → 결과 → 닫기 → **검색 재호출하여 자료수 0건 확인**. 0건 아니면 재시도.

### Step 3 — 사방넷 001→002 변환 (자동 ✅, popup VM + 0건 검증)
`#/order/order-confirm` — popup VM `sbForm.orderStatus='002'` 설정 + 저장 → confirm dialog 자동 클릭. **그 후 신규주문(001) 잔존 0건 확인**.

### Step 4 — list API 2종 join (자동 ✅)
- searchOrders(pageSize:1000, orderStatus:['002'], wyblNoregs:'Y')
- WaybillInputSku/getWaybillInputSkuLists(pageSize:1000, ordStsCd:'002', rmteArrdPrtYn:'Y')
- ordNo 키로 join

분류:
- 빈박스: 쿠팡 + 주소 `%` / 스마트스토어 + `shpmtMsg === '문 앞에 놓아주세요!'` 정확 매칭
- 04-28 이전 SKIP, state.json 차집합

### Step 5 — 풀필먼트 엑셀 생성 + 크로스체킹

```bash
python3 scripts/create_fulfillment_excel.py \
  --orders today_orders.json \
  --mapping scripts/product_mapping.json \
  --state <workspace>/realship_state.json \
  --output fulfillment_$(date +%Y%m%d).xlsx
```

또는 페이지 JS 매핑 패스 (이 SKILL의 룰 2번·15번 명세 그대로 사용 — Python과 동일 결과 보장).

**자동 sanity check (v0.2.6)**: `print_ffqty_distribution()` 호출 → ffQty 분포 출력 + 이상치 경고 + 사용자 confirm 항목 체크.

### Step 6 — 사업자별 분리 업로드

자동화 가능 범위:
- 이더 / 뉴트리 풀필먼트 wms02 로그인 ✅
- 발주등록 페이지 진입 ✅
- 엑셀등록 popup ✅
- file_upload tool 로 xlsx 업로드 ✅

**한계**: 등록 버튼 클릭 후 풀필먼트 응답 시간이 매우 김 (53건 처리 시 30초+ hung 가능). 사용자 직접 fallback 권장 또는 timeout polling 룰 적용.

**v0.2.5 등록 직전 가드**: list API 재호출 → ordStsCd ≠ '002' row 자동 SKIP + 사용자 confirm.

### Step 7 — state.json 갱신 + 종합 보고서 (자동 ✅)

```json
{
  "last_cycle_date": "2026-04-30",
  "2026-04-30": {
    "ether_ord_no_list": [...],
    "nutri_ord_no_list": [...],
    "shma_ord_no_list": [...],
    "cancelled_skipped": [{"ordNo": ..., "currentStatus": "..."}],
    "complete_time": "2026-04-30 14:50 KST"
  }
}
```

---

## 자동화 한계 (Chrome MCP 환경)

| 항목 | 한계 | 사용자 작업 |
|---|---|---|
| 풀필먼트 등록 버튼 click 이후 응답 | hung 또는 30초+ 대기 | 직접 click 후 모니터 또는 timeout polling |

(v0.2.5 이전의 "풀필먼트 사이트 진입" 항목은 v0.2.6에서 제거 — 자동 진입 가능 확인)

---

## 핵심 가드 (CRITICAL)

1. shmaOrdNo 단일 키 매칭
2. **수량 룰 v3** — set_multiplier(N+N → N) × cnt_multiplier(채널 분기)
3. 빈박스 `!` 정확 매칭
4. state.json 차집합
5. 헤더 HEADERS 상수 강제
6. 스마트스토어 일주일치 크로스체킹 의무
7. list API 2종 join — 다운로드 사용 금지
8. 등록 직전 ordStsCd 재확인
9. **🆕 단계마다 0건 검증 의무**
10. **🆕 Python ↔ JS 매핑 단일 명세**
11. **🆕 ffQty 분포 자동 출력 + 이상치 경고**

---

## 변경 이력

- **v0.2.6** (2026-04-30) — 수량 룰 v3 (cnt_multiplier 채널 분기 신규), 비밀번호 정정 (! → @), 시리즈 B (남녀공용 초경량 아쿠아슈즈 논슬립 워터슈즈, E00400656) 키워드 룰 추가, 단계마다 0건 검증 의무 룰, Python·JS 매핑 단일 명세 룰, SheetJS jsdelivr 우선, list API body sbForm deep copy 룰. 자동화 한계 표 갱신.
- **v0.2.5** (2026-04-30) — 엑셀 다운로드 폐지 (list API 2종 join), 등록 직전 ordStsCd 재확인 가드.
- **v0.2.4** (2026-04-29) — 송장처리 분리, 빈박스 룰 정정, 매핑 7종, state.json 차집합, 스마트스토어 크로스체킹.
- **v0.2.3** (2026-04-28) — 폴더 구조·schema 정정.
- **v0.2.2** (2026-04-28) — plugin-github-sync 스킬.
- **v0.2.1** (2026-04-28) — 중복 발주 가드.
- **v0.2.0** (2026-04-28) — 수량 룰 v2, shmaOrdNo 단일 키, 사전 일괄확정·001→002 의무화.
