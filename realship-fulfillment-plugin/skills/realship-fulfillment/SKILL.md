---
name: realship-fulfillment
description: >
  실배송 풀필먼트 자동 등록 스킬 (v0.2.4 — 송장처리 제외 / 풀필먼트 등록까지만).
  매일 오후 2시 주문 마감 후 사방넷 → 사방넷 풀필먼트(이더+뉴트리)에 실배송만 자동 등록.
  반드시 이 스킬을 사용해야 하는 경우: "실배송 풀필먼트", "풀필먼트 등록", "실배송 등록", "오늘 실배송",
  "풀필먼트 발주", "실배송 처리해줘", "주문 마감", "2시 마감", "풀필먼트에 넣어줘", "실배송 올려줘",
  "발주등록해줘", "실배송 1차", "실배송 2차", "사방넷 발주", "스마트스토어 발주", "발주 엑셀 만들어줘"
---

# 실배송 풀필먼트 자동 등록 (v0.2.4)

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
| 이더컴퍼니 (공산품) | `w7298` | `eithercompany` | `dlejrhddyd1!` | E-코드 |
| 뉴트리정 (영양제) | `w7298` | `nutrijung` | `dlejrhddyd1!` | N-코드 |

URL: `https://wms02.sbfulfillment.co.kr`

⚠️ **풀필먼트 다운로드 엑셀은 비밀번호 보호** — 위 비번으로 복호화 (msoffcrypto-tool)

---

## v0.2.4 절대 룰 (사고 기반 영구 반영)

1. **shmaOrdNo 단일 키 매칭** — 받는분 이름 매칭 절대 금지 (김채은 cross-account 사고 차단)
2. **수량 룰 v2** — `풀필 수량 = ordQty × set_multiplier(옵션)` (1+1/2+2/3+3/N개)
3. **빈박스 룰 정정** — `'문 앞에 놓아주세요!' in msg` 느낌표 정확 매칭만 빈박스. 부분 매칭 금지!
   - 04-29 사고: 부분 매칭으로 9건(스마트스토어 일반 안전 메시지) 누락 직전
4. **사전 일괄확정·001→002 자동 의무화** (Vue VM 트릭)
5. **풀필먼트 엑셀 생성 후 사용자 수량 검토 의무**
6. **04-28 같은 사이클 시작일 이전 주문 자동 SKIP** (수동처리분 중복 차단)
7. **state.json: 오늘 풀필먼트 등록 ord_no 저장** → 다음 사이클 차집합 가드 (어제 발주조회 다운 불필요)
8. **풀필먼트 발주 엑셀 헤더는 반드시 스크립트의 HEADERS 상수 import 사용** — 직접 작성 금지
9. **빈 전화/우편/주소 fallback** — 010-0000-0000 / 00000 / "주소 사방넷 자동입력"
10. **출고희망일 자동 = 다음날 (YYYY-MM-DD)**
11. **풀필먼트 업로드 전 스마트스토어 일주일치 vs 풀필먼트 엑셀 크로스체킹 의무** — 14:00 이후 신규 / 빈박스 잘못 분류 / history 잘못 제외 모두 캐치

---

## 매 사이클 표준 6단계 (송장처리 제외)

### Step 1 — 사방넷 주문수집 (자동 ✅)

`https://sbadmin03.sabangnet.co.kr/#/order/order-collect-auto`

```javascript
// el-table에서 7개 mall row 직접 push (toggleAllSelection은 안 먹음)
const tables = document.querySelectorAll('.el-table');
let target = null;
tables.forEach(t => { const v = t.__vue__; if (v && v.store && v.store.states.data.length === 7) target = v; });
const data = target.store.states.data;
data.forEach(r => target.toggleRowSelection(r, true));
```

→ "주문수집(신규+주문확인)" 버튼 좌표 (763, 353) 클릭 → 확인 모달 (880, 517) 클릭

→ 새 탭(`127.0.0.1:8181/mall_join/auto_service/client_order_collect.html`) 열리며 SabangSCM 자동 처리

### Step 2 — 사방넷 일괄확정 (자동 ✅)

`https://sbadmin03.sabangnet.co.kr/#/order/order-decide`

→ 검색 (935, 252) → 일괄주문확정 (1308, 290) → 일괄주문확정 모달 (768, 410) → 확인 (880, 386)

### Step 3 — 사방넷 001→002 변환 (자동 ✅, Vue VM 트릭)

`#/order/order-confirm` 페이지에서 popup VM 직접 호출 (SKILL 기존 패턴 유지):
```javascript
window.sbParamMap['order-confirm-order-status-change-popup'] = { bindObject: { dataList, ordNoArr } };
Object.defineProperty(window, 'opener', { get: () => window });
window.name = 'order-confirm-order-status-change-popup';
window.location.hash = '#/popup/...';
// popup VM 잡고 → exeOrderConfirmOrderStatusChange()
```

### Step 4 — 주문서확인 엑셀 다운로드 (⚠️ 사용자 직접 필요)

**한계**: `vm.makeExcelDownload()` 직접 호출 시 400 에러 (body schema 미지). 화면 클릭은 60초 락 + popup 차단으로 무반응.

**워크어라운드**: 사용자가 30초 안에 직접:
1. 주문서확인처리 → 송장미등록 체크 → 검색
2. 양식 dropdown → 양식1 / 전체자료
3. 다운로드 → 엑셀 받기 → 채팅 업로드

스크립트가 받자마자 분류·매핑 자동:
- 빈박스 SKIP: 쿠팡 + 주소 `%` / 스마트스토어 + `'문 앞에 놓아주세요!'` 정확 매칭
- 04-28 이전 주문일자 SKIP (수동처리분)
- state.json의 `last_fulfillment_ord_no` 차집합 (다음 사이클 중복 차단)
- 실배송 추출 → product_mapping.json 매핑

### Step 5 — 풀필먼트 엑셀 생성 + 크로스체킹

```bash
python3 scripts/create_fulfillment_excel.py \
  --orders today_orders.json \
  --mapping scripts/product_mapping.json \
  --state <workspace>/realship_state.json \
  --output fulfillment_$(date +%Y%m%d).xlsx
```

**필수 검증 (사용자 confirm 의무)**:
1. 사업자별 분리 결과 (이더 E* / 뉴트리 N* / 미분류)
2. 복수구매 그룹 (같은 받는분+주소) → 합배송 OK 여부
3. 동성동명 의심 (같은 이름, 다른 주소) → 별도 발송 OK 여부
4. 미분류 → product_mapping.json 키워드 룰 추가 또는 사용자 확정 코드

**스마트스토어 일주일치 크로스체킹 의무 (CRITICAL)**:
사용자에게 스마트스토어 주문조회 일주일치 엑셀 요청 → 우리 엑셀 매칭:
- **누락 위험**: 스마트스토어 발송대기인데 우리 엑셀에 없음 → 빈박스 잘못 분류 또는 매핑 누락 또는 14:00 이후 신규
- **중복 위험**: 우리 엑셀에 있는데 스마트스토어 이미 발송완료 → 사방넷 history 잘못 제외 (옵션 다른 추가주문)

### Step 6 — 사업자별 분리 업로드 (⚠️ 사용자 직접)

**한계**: `wms02.sbfulfillment.co.kr` 페이지가 Chrome MCP 환경에서 spinner만 회전 (Vue 3 hydration 차단)

**워크어라운드**: 사용자가 직접:
1. 풀필먼트 로그인 (이더 → 뉴트리 두 계정)
2. 발주등록 → 엑셀 업로드
3. 4건 이상 주소·전화 placeholder인 어제 누락분은 화면에서 직접 수정

### Step 7 — state.json 갱신 + 종합 보고서 (자동 ✅)

```json
{
  "last_cycle_date": "2026-04-29",
  "2026-04-29": {
    "ether_ord_no_list": [...],
    "nutri_ord_no_list": [...],
    "shma_ord_no_list": [...],
    "complete_time": "2026-04-29 14:50 KST"
  }
}
```

**state 위치**: 사용자 워크스페이스 폴더 (예: `<Downloads>/realship_state.json`). 플러그인 폴더에는 저장 금지 (read-only).

종합 보고서 HTML 자동 생성 → 처리 결과 + 발견된 사고 + 매핑 학습 + v다음 패치 항목

---

## 자동화 한계 (Chrome MCP 환경 - 사용자 직접 필요)

| 항목 | 한계 | 사용자 작업 |
|---|---|---|
| 사방넷 주문서확인처리 엑셀 다운 | API 400 / popup 차단 | 30초 (양식1 + 전체자료 + 다운) |
| 풀필먼트 사이트 진입 | Vue 3 hydration 차단 | 풀필먼트 로그인 + 발주등록 (5분) |
| 풀필먼트 엑셀 다운 (state 갱신용) | 동일 | (제거됨 - state.json으로 대체) |

---

## 핵심 가드 (CRITICAL)

1. **shmaOrdNo 단일 키 매칭** (받는분 이름 매칭 절대 금지)
2. **수량 룰 v2** — set_multiplier(옵션) 적용
3. **04-29 빈박스 룰 정정** — `'문 앞에 놓아주세요!' in msg` 느낌표 정확 매칭만
4. **state.json 차집합** — `same shmaOrdNo + same 사방넷코드` 모두 매칭 시에만 SKIP
5. **헤더는 스크립트 HEADERS 상수 강제** — 발주 엑셀 직접 작성 금지
6. **스마트스토어 일주일치 크로스체킹 의무** — 누락/중복 모두 캐치

---

## 변경 이력

- **v0.2.4** (2026-04-29) — 송장처리 제거 (waybill-processing 별도 스킬 분리), 빈박스 룰 정정 (`!` 정확 매칭), 매핑 룰 7종 추가, state.json으로 풀필먼트 발주조회 다운 대체, 스마트스토어 크로스체킹 의무 룰, 자동화 한계 명시.
- **v0.2.3** (2026-04-28) — 폴더 구조 정정, plugin/marketplace 스키마 정정.
- **v0.2.2** (2026-04-28) — plugin-github-sync 스킬 추가.
- **v0.2.1** (2026-04-28) — 중복 발주 가드 (`--history`).
- **v0.2.0** (2026-04-28) — 수량 룰 v2, shmaOrdNo 단일 키, 사전 일괄확정·001→002 의무화.

