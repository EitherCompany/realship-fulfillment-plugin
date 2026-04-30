---
name: realship-fulfillment
description: >
  실배송 풀필먼트 자동 등록 스킬 (v0.2.5 — 엑셀 다운로드 완전 폐지 / list API 2종 join + 취소 가드).
  매일 오후 2시 주문 마감 후 사방넷 → 사방넷 풀필먼트(이더+뉴트리)에 실배송만 자동 등록.
  반드시 이 스킬을 사용해야 하는 경우: "실배송 풀필먼트", "풀필먼트 등록", "실배송 등록",
  "오늘 실배송", "풀필먼트 발주", "실배송 처리해줘", "주문 마감", "2시 마감",
  "풀필먼트에 넣어줘", "실배송 올려줘", "발주등록해줘", "실배송 1차", "실배송 2차",
  "사방넷 발주", "스마트스토어 발주", "발주 엑셀 만들어줘"
---

# 실배송 풀필먼트 자동 등록 (v0.2.5)

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

## v0.2.5 절대 룰 (사고 기반 영구 반영)

1. **shmaOrdNo 단일 키 매칭** — 받는분 이름 매칭 절대 금지 (김채은 cross-account 사고 차단)
2. **수량 룰 v2** — `풀필 수량 = ordQty × set_multiplier(옵션)` (1+1/2+2/3+3/N개)
3. **빈박스 룰 정정** — `'문 앞에 놓아주세요!' in shpmtMsg` 느낌표 정확 매칭만 빈박스. 부분 매칭 금지!
   - 04-29 사고: 부분 매칭으로 9건(스마트스토어 일반 안전 메시지) 누락 직전
4. **사전 일괄확정·001→002 자동 의무화** (Vue VM 트릭)
5. **풀필먼트 엑셀 생성 후 사용자 수량 검토 의무**
6. **04-28 같은 사이클 시작일 이전 주문 자동 SKIP** (수동처리분 중복 차단)
7. **state.json: 오늘 풀필먼트 등록 ord_no 저장** → 다음 사이클 차집합 가드
8. **풀필먼트 발주 엑셀 헤더는 반드시 스크립트의 HEADERS 상수 import 사용** — 직접 작성 금지
9. **빈 전화/우편/주소 fallback** — 010-0000-0000 / 00000 / "주소 사방넷 자동입력"
10. **출고희망일 자동 = 다음날 (YYYY-MM-DD)**
11. **스마트스토어 일주일치 vs 풀필먼트 엑셀 크로스체킹 의무** (14:00 이후 신규 / 빈박스 잘못 분류 / history 잘못 제외 모두 캐치)
12. **🆕 list API 2종 join 패턴 — 엑셀 다운로드 완전 폐지** (v0.2.5 검증 완료)
13. **🆕 풀필먼트 등록 직전 ordStsCd 재확인 가드 의무** — 취소·반품 사고 차단

---

## v0.2.5 핵심 우회 (2026-04-30 검증 완료)

**문제**: 사방넷 엑셀 다운로드는 popup 차단·chunk 한계·인증 헤더 등으로 자동화 어려움. v0.2.4까지 사용자 직접 다운로드가 한계 항목이었음.

**해결**: 두 list API를 직접 호출해 ordNo 키로 join하면 다운로드 엑셀과 동등한 데이터셋 확보 가능.

| API | 응답에 있는 핵심 필드 | 응답에 없는 필드 |
|---|---|---|
| `/prod-api/customer/order/OrderConfirm/searchOrders` | `shpmtMsg`, `shpmtEtcFldVl`, `shmaNm`, `ordClctFldVl1~4` | `ecptRmteTotAddr`, `rmteZipcd` |
| `/prod-api/customer/order/WaybillInputSku/getWaybillInputSkuLists` | `ecptRmteTotAddr`, `rmteZipcd`, `ecptRmteNm`, `ecptRmteTelNo`, `ecptRmteHndpnNo`, `clctPrdNm`, `clctSkuNm`, `ordQt`, `ordStsCd` | `shpmtMsg` |

**ordNo 키로 join → 100% 완전 데이터셋.** 엑셀 다운로드 불필요.

### 인증 (cookie + Bearer)
페이지에서 사용 중인 axios가 추가하는 헤더 그대로 사용:
```javascript
// XHR setRequestHeader 가로채서 Authorization Bearer JWT 캡쳐
// 또는 document.cookie의 Authorization 값 사용
const auth = 'Bearer ' + getCookieValue('Authorization');  // 또는 캡쳐된 JWT
```

### 호출 패턴
```javascript
// 1. searchOrders — shpmtMsg
const sBody = {
  regDmStartTime:'00:00', regDmEndTime:'23:00', fnlChgPrgmNm:'order-confirm',
  chkOrdNo:[], checkList:[], currentPage:1,
  dateDiv:'reg_dm', startDate:'YYYYMMDD', endDate:'YYYYMMDD',
  pageSize:1000, orderStrd:'fst_regs_dt', orderDegreeStrd:'desc',
  wyblNoregs:'Y', orderStatus:['002'],
  prdNmDiv:'prod_nm', shopType:'mall',
  searchCondition:'cust_nm', searchConditionOption:'p_sku_value', searchKeywordList:[],
  svcAcntId:'mw159514', userId:'eithercompany',
  // ... 나머지 default 값들 (캡쳐된 body 그대로)
};
const sRes = await fetch('/prod-api/customer/order/OrderConfirm/searchOrders', {
  method:'POST', body:JSON.stringify(sBody), credentials:'include',
  headers:{'Content-Type':'application/json', 'Authorization':auth}
}).then(r => r.json());
const msgMap = {}; sRes.data.orderList.forEach(r => { msgMap[r.ordNo] = {shpmtMsg: r.shpmtMsg, shmaNm: r.shmaNm, shmaCnctnLoginId: r.shmaCnctnLoginId}; });

// 2. WaybillInputSku — addr/zip
const wBody = {/* 캡쳐된 body 그대로 + pageSize:1000, ordStsCd:'002', startDate, endDate */};
const wRes = await fetch('/prod-api/customer/order/WaybillInputSku/getWaybillInputSkuLists', {
  method:'POST', body:JSON.stringify(wBody), credentials:'include',
  headers:{'Content-Type':'application/json', 'Authorization':auth}
}).then(r => r.json());

// 3. join (ordNo 키)
const joined = wRes.data.list.filter(r => r.ordNo).map(w => ({...w, ...msgMap[w.ordNo]}));
```

---

## 매 사이클 표준 6단계

### Step 1 — 사방넷 주문수집 (자동 ✅)
`https://sbadmin03.sabangnet.co.kr/#/order/order-collect-auto` — 7개 mall row 일괄 선택 후 `주문수집(신규+주문확인)` → 확인.

### Step 2 — 사방넷 일괄확정 (자동 ✅)
`#/order/order-decide` — 검색 → 일괄주문확정 → 확인.

### Step 3 — 사방넷 001→002 변환 (자동 ✅, popup VM 트릭)
`#/order/order-confirm` — `window.sbPopupMap['order-confirm-order-status-change-popup']` 직접 참조 → popup vm의 `sbForm.orderStatus = '002'` 설정 → 저장 버튼 click → confirm dialog 자동 클릭.

### Step 4 — list API 2종 join (자동 ✅, 신규 v0.2.5)
**엑셀 다운로드 폐지.** 위 호출 패턴으로 두 API 호출 + ordNo join → 73건 데이터 확보.

분류 룰 적용:
- 빈박스 SKIP: 쿠팡 + 주소 `%` / 스마트스토어 + `shpmtMsg === '문 앞에 놓아주세요!'` 정확 매칭
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

또는 페이지 안에서 SheetJS로 직접 생성 → blob 자동 다운로드.

**필수 검증 (사용자 confirm 의무)**:
1. 사업자별 분리 결과 (이더 E* / 뉴트리 N* / 미분류)
2. 복수구매 그룹 (같은 받는분+주소) → 합배송 OK 여부
3. 동성동명 의심 (같은 이름, 다른 주소) → 별도 발송 OK 여부
4. 미분류 → product_mapping.json 키워드 룰 추가 또는 사용자 확정 코드

**스마트스토어 일주일치 크로스체킹 의무 (CRITICAL)**

### Step 6 — 사업자별 분리 업로드 (사용자 직접 + 등록 직전 취소 가드)

**🆕 v0.2.5 취소 가드 — 풀필먼트 엑셀 업로드 직전 의무**

```javascript
// 등록 직전 list API 재호출 → 현재 ordStsCd 002가 아닌 row 자동 SKIP
const recheckBody = {/* 같은 body, ordStsCd 필터 없이 */};
const recheck = await fetch('/prod-api/customer/order/WaybillInputSku/getWaybillInputSkuLists', {...}).then(r => r.json());
const currentStatus = {};
recheck.data.list.filter(r => r.ordNo).forEach(r => { currentStatus[r.ordNo] = r.ordStsCd; });

// 풀필먼트 엑셀의 각 row에 대해
const safeRows = [];
const cancelledRows = [];
for (const row of fulfillmentRows) {
  const cur = currentStatus[row.ordNo];
  if (cur === '002') safeRows.push(row);
  else cancelledRows.push({...row, currentStatus: cur});
}

// 취소 의심건은 사용자 confirm 받기 의무
if (cancelledRows.length > 0) {
  // 사용자에게 표시: "{ordNo} {shmaOrdNo} 가 현재 {cur} 상태로 변경됨. 풀필먼트 등록 SKIP 권장."
}
```

**왜**: 주문수집 후 사용자가 주문을 취소·반품했는데 우리 풀필먼트 엑셀에는 등록되는 사고가 발생했음. 등록 직전 재확인이 의무.

업로드 절차:
1. 풀필먼트 로그인 (이더 → 뉴트리 두 계정)
2. 발주등록 → 풀필먼트 엑셀(취소 가드 적용 후) 업로드
3. 4건 이상 주소·전화 placeholder인 어제 누락분은 화면에서 직접 수정

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

종합 보고서 HTML 자동 생성.

---

## 자동화 한계 (Chrome MCP 환경)

| 항목 | 한계 | 사용자 작업 |
|---|---|---|
| ~~사방넷 주문서확인 엑셀 다운~~ | **v0.2.5에서 list API 2종 join으로 폐지** | (제거됨) |
| 풀필먼트 사이트 진입 | Vue 3 hydration 차단 | 풀필먼트 로그인 + 발주등록 (5분) |

---

## 핵심 가드 (CRITICAL)

1. **shmaOrdNo 단일 키 매칭** (받는분 이름 매칭 절대 금지)
2. **수량 룰 v2** — set_multiplier(옵션) 적용
3. **빈박스 룰 정정** — `'문 앞에 놓아주세요!' in shpmtMsg` 느낌표 정확 매칭만
4. **state.json 차집합** — `same shmaOrdNo + same 사방넷코드` 모두 매칭 시에만 SKIP
5. **헤더는 스크립트 HEADERS 상수 강제** — 발주 엑셀 직접 작성 금지
6. **스마트스토어 일주일치 크로스체킹 의무** — 누락/중복 모두 캐치
7. **🆕 list API 2종 join — 엑셀 다운로드 사용 금지** (v0.2.5)
8. **🆕 등록 직전 ordStsCd 재확인 — 취소 row 자동 SKIP** (v0.2.5)

---

## 변경 이력

- **v0.2.5** (2026-04-30) — 엑셀 다운로드 완전 폐지. searchOrders + WaybillInputSku list API 2종 join 패턴 도입. 등록 직전 ordStsCd 재확인 가드 추가 (취소·반품 row 자동 SKIP). 자동화 한계 항목 1개 제거.
- **v0.2.4** (2026-04-29) — 송장처리 제거 (waybill-processing 별도 스킬 분리), 빈박스 룰 정정 (`!` 정확 매칭), 매핑 룰 7종 추가, state.json으로 풀필먼트 발주조회 다운 대체, 스마트스토어 크로스체킹 의무 룰, 자동화 한계 명시.
- **v0.2.3** (2026-04-28) — 폴더 구조 정정, plugin/marketplace 스키마 정정.
- **v0.2.2** (2026-04-28) — plugin-github-sync 스킬 추가.
- **v0.2.1** (2026-04-28) — 중복 발주 가드 (`--history`).
- **v0.2.0** (2026-04-28) — 수량 룰 v2, shmaOrdNo 단일 키, 사전 일괄확정·001→002 의무화.
