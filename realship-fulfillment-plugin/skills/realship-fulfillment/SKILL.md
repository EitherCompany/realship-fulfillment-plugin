---
name: realship-fulfillment
description: >
  실배송 주문 풀필먼트 자동 등록 + 송장처리 통합 스킬 (v0.2.3).
  매일 1차/2차 사이클로 사방넷 → 사방넷 풀필먼트(이더+뉴트리)에 실배송만 자동 등록.
  핵심 가드: shmaOrdNo 단일 키 매칭 (받는분 이름 절대 사용 금지),
  ordQty × set_multiplier 수량 룰, 직전 7일 풀필 history 기반 중복 발주 차단,
  사전 일괄확정·001→002 상태변경 의무화, 빈박스 후보 보수적 SKIP.
  반드시 이 스킬을 사용해야 하는 경우: "실배송 풀필먼트", "풀필먼트 등록", "실배송 등록",
  "오늘 실배송", "풀필먼트 발주", "실배송 처리해줘", "주문 마감", "2시 마감",
  "풀필먼트에 넣어줘", "실배송 올려줘", "발주등록해줘", "실배송 1차", "실배송 2차",
  "송장처리", "운송장 처리", "쇼핑몰 송신", "스마트스토어 발주", "사방넷 발주",
  "발주 엑셀 만들어줘" 등 실배송 주문을 사방넷/풀필먼트에 등록·송장처리하는 모든 요청.
---

# 실배송 풀필먼트 자동 등록 + 송장처리 (v0.2.3)

## 시스템 정보

### 사방넷 관리자
| 항목 | 값 |
|---|---|
| URL | `https://sbadmin03.sabangnet.co.kr` |
| 로그인 | `eithercompany` / `dlejzja7801!` |
| svcAcntId | `mw159514` |

### 사방넷 풀필먼트 (계정 2개)
| 계정 | 회사코드 | 아이디 | 비밀번호 | 취급 |
|---|---|---|---|---|
| 이더컴퍼니 (공산품) | `w7298` | `eithercompany` | `dlejrhddyd1@` | E-코드 |
| 뉴트리정 (영양제) | `w7298` | `nutrijung` | `dlejrhddyd1@` | N-코드 |

URL: `https://wms02.sbfulfillment.co.kr`

⚠️ 비밀번호 5회 실패 시 10분 잠김. 엉뚱한 계정 업로드 시 "신규 상품 등록" 모달 → **무조건 [취소]**.

---

## v0.2.3 절대 룰 (창근님 결정 사항 영구 반영)

1. **shmaOrdNo 단일 키 매칭** — 받는분 이름 매칭 절대 금지. 김채은 동성동명 cross-account 사고 차단.
2. **수량 룰 v2** — `풀필 수량 = ordQty × set_multiplier(옵션)`. 김현숙시티(2→1)·홍주미(2→1)·권나영(3→1) 누락 사고 차단.
3. **중복 발주 가드** — 직전 7일 풀필 history 비교. 같은 쇼핑몰주문번호 발견 시 exit 2 차단.
4. **사전 일괄확정·001→002 자동** — Vue VM 직접 호출 트릭으로 매번 자동 (사용자 클릭 요청 금지).
5. **풀필먼트 등록 엑셀 생성 후 → 사용자 수량 검토 의무**.
6. **빈박스 후보 단순 룰** — "문 앞에 놓아주세요!" → SKIP, 다음 사이클에서 또 SKIP되면 사용자 확인.
7. **송장처리 매칭 = shmaOrdNo 단일 키**.
8. **사이클 마지막에 종합 보고서 자동 생성**.
9. **처리 범위 = `last_fulfillment_upload_date` 14:01 ~ 어제 14:00** (1일 lag, 빈박스 명단 도착 보장).

---

## 매 사이클 표준 8단계

### Step 0 — 풀필먼트 발주조회 raw 다운로드 (중복 가드용)

풀필먼트 (이더+뉴트리) 두 계정 모두 → 발주조회 → 직전 7일 → "쇼핑몰주문번호" 컬럼 체크 → 엑셀 다운 → JSON 변환.

### Step 1 — 사방넷 주문수집

주문서수집(자동) 결과 확인 → 5분 대기 후 재조회 → 0건이면 사용자 보고.

### Step 2 — 사방넷 일괄확정 (주문서확정관리, 자동)

7일 범위 + "주문미확정" 필터 → 일괄주문확정. ordCnfrmYn N → Y.

### Step 3 — 사방넷 주문상태변경 001→002 (Vue VM 직접 호출)

```javascript
// 주문서확인처리 페이지에서:
const root = document.getElementById('app');
const inst = root.__vue__;

// 1) 신규주문(001) 데이터 선택
let tab = null;
const findTab = (i) => {
  if (i?.$options?.name === 'ElTable' && i.store.states.data.length > 100) { tab = i; return; }
  i?.$children?.forEach(findTab);
};
findTab(inst);
const targetRows = tab.store.states.data.filter(r => r.ordStsCd === '001');
tab.clearSelection();
targetRows.forEach(r => tab.toggleRowSelection(r, true));

// 2) sbParamMap 주입 + window.opener=window + name 설정
window.sbParamMap = window.sbParamMap || {};
window.sbParamMap['order-confirm-order-status-change-popup'] = {
  bindObject: { dataList: targetRows, ordNoArr: targetRows.map(r => r.ordNo) },
  resultFn: () => {}
};
Object.defineProperty(window, 'opener', { get: () => window, configurable: true });
window.name = 'order-confirm-order-status-change-popup';

// 3) popup URL 로 navigate (같은 탭)
window.location.hash = '#/popup/views/pages/order/order-confirm/order-confirm-order-status-change-popup.vue?menuNo=661';

// 4) popup VM 찾고 데이터 set 후 직접 호출
let popVm = null;
const findPopVm = (i) => {
  if (i?.$data?.tableData !== undefined && i?.$data?.modifyAllList !== undefined) { popVm = i; return; }
  i?.$children?.forEach(findPopVm);
};
findPopVm(document.getElementById('app').__vue__);

popVm.multiselectList.splice(0, 0, ...targetRows);
popVm.allList.splice(0, 0, ...targetRows);
popVm.tableData.splice(0, 0, ...targetRows);
popVm.sbForm.list = targetRows.slice();
popVm.sbForm.orderStatus = '002';
popVm.sbForm.selectData = '1';
popVm.compareChangeOrder = '1';
popVm.selectListSize = targetRows.length;
popVm.searchListSize = targetRows.length;

popVm.changeOrderStatus();
popVm.exeOrderConfirmOrderStatusChange();
```

API: `POST /prod-api/customer/order/OrderConfirm/exeOrderConfirmOrderStatusChange` → 200.

### Step 4 — 주문서확인 엑셀 다운로드 + 분류 + 매핑

분류:
- 쿠팡 + 주소 `%` → 빈박스 SKIP
- "문 앞에 놓아주세요!" → 빈박스 후보 SKIP (state.json 기록)
- 나머지 → 실배송

차집합 적용 (1차: 어제까지 등록 ord 제외, 2차: 1차 등록 ord 제외).

### Step 5 — 풀필먼트 엑셀 생성 (가드 자동)

```bash
python3 scripts/create_fulfillment_excel.py \
  --orders today_orders.json \
  --mapping scripts/product_mapping.json \
  --history fulfillment_history_*.json \
  --output fulfillment_$(date +%Y%m%d)_1차.xlsx
```

3대 가드: 중복 발주(v0.2.1), 수량 누락(v2), 고수량 경고.

🚨 **사용자 수량 검토 의무**: 엑셀 생성 후 수량 표 제출 → 승인 후 다음 단계.

### Step 6 — 사업자별 분리 업로드

자동 분리:
- `_ether.xlsx` → eithercompany 계정
- `_nutri.xlsx` → nutrijung 계정
- `_unmapped.xlsx` → 사용자 수동

⚠️ freeze 패턴: 30초 응답 없으면 25~30초 대기 → `/order/add` 재진입 → 업로드 이력 확인.

### Step 7 — state.json 갱신

```json
{
  "last_fulfillment_upload_date": "2026-04-29",
  "2026-04-29": {
    "1차_실배송_사방넷_ord_no": [...],
    "1차_풀필먼트_등록_쇼핑몰주문번호": [...],
    "빈박스_후보_사방넷_ord_no": [...],
    "1차_완료시각": "2026-04-29 11:30 KST"
  }
}
```

### Step 8 — 종합 보고서

HTML 리포트 자동 생성, 사용자에게 링크 제출.

---

## 송장처리 사이클 (1차+2차 일괄, 보통 풀필 송장 발행 후)

1. 풀필먼트 발주조회 → 오늘 발행 송장 추출
2. `match_waybills.py` shmaOrdNo 단일 키 매칭
3. 사방넷 운송장입력(대량) — el-upload `handleStart(file)` 직접 호출
4. 쇼핑몰운송장송신 — window.open fake hook + `sendWybl()`
5. mall 거부 3종 분류 (취소거부 후 수동 / 이미 취소 / cross-account)

---

## 빈박스 스킬과의 연동

- 두 스킬 모두 사방넷 일괄확정·상태변경 수행 → state.json 의 플래그로 중복 작업 방지
- 빈박스 사이클이 처리한 ord_no 는 `빈박스_처리완료_사방넷_ord_no` 에 기록 → 실배송이 차집합으로 제외
- 사이클 순서 자유 — 단 실배송에서 "문 앞 배송" 메시지 ord 는 보수적 SKIP, 다음 사이클 재확인

---

## 변경 이력

- **v0.2.3** (2026-04-28) — 마켓플레이스 스키마 정정 (binbox 패턴 정확 복제). 폴더 구조: `realship-fulfillment-plugin/` 하위 폴더로 이동. marketplace.json `$schema` + `owner` + `plugins[]` 배열. plugin.json `keywords` + `homepage` + `repository` + `license:MIT`.
- **v0.2.2** (2026-04-28) — 단순 룰 정리, plugin-github-sync 스킬 추가, 1차/2차 분할 사이클 명시.
- **v0.2.1** (2026-04-28) — 중복 발주 가드 (`--history`), `last_fulfillment_upload_date` 추가.
- **v0.2.0** (2026-04-28) — 수량 룰 v2, shmaOrdNo 단일 키, 사전 일괄확정·001→002 의무화.
- **v0.1.0** (2026-04-28) — 초기 패키징.

---

## 핵심 학습 7가지 (2026-04-27 + 04-28 사고 기반)

1. ordQty 무시 → v2 가드
2. 받는분 이름 매칭 → shmaOrdNo 단일 키
3. 후행 신규주문 누락 → 사전 일괄확정·001→002 의무화
4. 주문상태변경 popup 차단 → window.opener=window + sbParamMap
5. 쇼핑몰송신 popup 차단 → window.open fake 객체
6. 9건 중복 발송 → 직전 7일 history 가드
7. mall 거부 3종 분류 (취소거부 후 수동 / 이미 취소 / cross-account)
