# 내일 실행 플로우 체크리스트 (2026-04-24)

오늘(2026-04-23) 실전에서 발견된 이슈를 모두 반영한 내일 실행 가이드.

## ⏰ 타임라인

| 시간 | 작업 | 스킬 |
|------|------|------|
| ~11:00 | 1차 실배송 테스트 | realship-fulfillment |
| 14:00 | 주문 마감 | - |
| 14:00~14:30 | 2차 실배송 (본 실행) | realship-fulfillment |
| 이더/뉴트리 풀필먼트 업로드 후 | 사방넷 주문확정 (001→002) | realship-fulfillment (확장) |
| 풀필먼트 송장 배정 후 | 송장 처리 + 송신 + 리포트 | waybill-processing |

---

## 🚨 오늘 사고 기반 필수 주의사항

### 1. **주문확정(001→002)은 운송장입력 前 선행 필수**
- 증상: 주문 상태가 001(신규주문)이면 운송장 대량입력 시 "송장을 등록할 주문상태가 아닙니다" 실패
- 오늘 결과: 73건 중 57건 성공 / 16건 실패 (모두 001 상태)
- **대응**:
  - `updateOrdStsCd` API 500 에러 3일째 (04-20/21/23 모두 실패) → 사용하지 말 것
  - 사방넷 UI **"주문서확정관리 → 주문미확정 검색 → 일괄주문확정"** 경로가 가장 안정적
  - 주문서확인처리에서 상태필터 001 → 전체선택 → "주문상태변경" 버튼도 가능

### 2. **쇼핑몰 송신은 쿠팡 vs 스마트스토어 방식이 다름**
| 쇼핑몰 | 방식 |
|--------|------|
| 쿠팡(shop0075) | 단순 `운송장송신` 버튼 → 자동 API 처리 |
| 스마트스토어(shop0055) | **4단계 패턴 필수** (iframe 127.0.0.1:8181 경유) |

### 3. **4단계 패턴 — 검증된 유일한 스마트스토어 송신 방법**
```javascript
// Step 1: getWaybillTransmitInfo
xhr.open('POST', '.../prod-api/customer/mall/MallWaybillTransmit/getWaybillTransmitInfo');
xhr.send(JSON.stringify({ svcAcntId: 'mw159514', ordNoList: [...] }));
// → window.__sendDatas = response.data.list

// Step 2: 모달 오픈
comp.mallWaybillTransmitPopup(window.__sendDatas, 'N');

// Step 3: iframe name 세팅 (다음 javascript_tool 호출에서)
document.querySelector('iframe[src*="127.0.0.1:8181"]').name = 'mallWayBillSong';

// Step 4: form submit
document.querySelector('form[target="mallWayBillSong"]').submit();
// → 10초 대기 후 결과 확인
```

### 4. **Vue 테이블 체크박스는 DOM 클릭으로 안 됨**
- DOM의 `label.el-checkbox.is-checked` 상태는 바뀌지만 Vue store.selection은 비어있음
- **해결**: 
```javascript
const el = document.querySelectorAll('.el-table')[0];
let vm = el.__vue__ || [findup __vue__];
vm.store.states.data.forEach(row => vm.toggleRowSelection(row, true));
// 이러면 vm.store.states.selection이 정상 업데이트
```

### 5. **4단계 패턴 재사용 시 iframe 초기화**
- 같은 페이지에서 두 번째 이후 송신 시도 시 iframe 캐시로 실패
- **해결**: 
```javascript
window.__cleanIframe = null;
location.reload();
// 또는 페이지 이동 후 재탐색
```

### 6. **택배사 코드**
| 택배사 | 사방넷 코드 | 사방넷 UI 표기 |
|-------|-----------|--------------|
| CJ대한통운 | 003 | CJ대한통운 |
| 한진택배 | 004 | 한진택배 |
| **롯데(현대)** | **005** | 롯데(현대)택배 / KGB택배 |

> 풀필먼트에서 "롯데택배"로 배정되어도 사방넷 UI는 "KGB택배"로 표시됨. 코드는 동일 005.

---

## ✅ 내일 1차 실배송 체크리스트 (11:00 경)

### Step 0: 준비
- [ ] Chrome MCP에서 사방넷 관리자(sbadmin03) 로그인 상태 확인
- [ ] 사방넷 풀필먼트(wms02) 양쪽 계정 확인:
  - 이더컴퍼니: `w7298 / eithercompany / dlejrhddyd1!`
  - 뉴트리정: `w7298 / nutrijung / dlejrhddyd1!`

### Step 1: 사방넷 주문수집 실행
- [ ] 사방넷 주문서수집(자동) 페이지 → 쇼핑몰 전체선택 → **"주문수집(신규+주문확인)"**
- [ ] 2-3분 대기 후 수집 상태 "정상종료" 확인
- ⚠️ 이 단계를 건너뛰면 최신 주문이 반영 안 됨

### Step 2: 주문서확인처리 엑셀 다운로드
- [ ] 일자 기준: **수집일** (배송희망일 아님!)
- [ ] 시작~종료: 어제 `00:00` ~ 오늘 `11:00`
- [ ] 다운로드 범위: **전체자료** (선택내용만 아님!)
- [ ] 양식: "테스트"
- [ ] 다운로드 → `C:\Users\Home\Downloads\20260424_주문서확인처리_테스트.xlsx`

### Step 3: 분류 + 매핑 + 엑셀 생성 (run_tomorrow.py)
```bash
python3 run_tomorrow.py "C:\Users\Home\Downloads\20260424_주문서확인처리_테스트.xlsx" --end-time "2026-04-24 11:00"
python3 convert_classified_to_sabang.py classified.json > sabang_orders.json
python3 create_fulfillment_excel.py --orders sabang_orders.json --mapping product_mapping.json --output fulfillment_20260424.xlsx
```

### Step 4: 풀필먼트 업로드 (이더 → 뉴트리)
- [ ] 이더컴퍼니 계정 로그인 → 발주등록 → 엑셀등록 → ether.xlsx
- [ ] 발주 결과 팝업 확인 (성공/실패 숫자)
- [ ] 로그아웃 → 뉴트리정 계정 로그인 → 발주등록 → 엑셀등록 → nutri.xlsx

### Step 5: **주문확정 (001→002) — 새로 추가된 단계**
- [ ] 사방넷 관리자 → 주문서확정관리
- [ ] 수집일자: 오늘 범위
- [ ] **주문미확정** 라디오 선택
- [ ] 검색 → 결과 건수가 풀필먼트 업로드 건수와 일치하는지 확인
- [ ] 전체선택 → **"일괄주문확정"** 클릭

### Step 6: 1차 크로스체킹 리포트
- [ ] 업로드 건수 / 매핑 누락 건수 / 쇼핑몰별 분포
- [ ] 누락 건 있으면 사용자에게 바로 보고

---

## ✅ 내일 2차 실배송 체크리스트 (14:00 경)

Step 1~6 동일하지만:
- [ ] 시간 필터 종료시각: `14:00` (또는 `14:30` 등 사용자 지정)
- [ ] 1차에서 이미 처리된 건은 `주문서확인처리`에 **안 뜸** (이미 002 이상 상태)
- [ ] 자동으로 신규 건만 처리됨

---

## ✅ 송장 처리 체크리스트 (waybill-processing)

### Step 1: 풀필먼트 송장 추출
- [ ] 이더컴퍼니 계정 → 발주조회 → 날짜 오늘 → 검색
- [ ] 테이블 DOM에서 `주문번호 / 택배사 / 송장번호` 추출 → `ether_waybills.json`
- [ ] 로그아웃 → 뉴트리정 계정 → 동일 추출 → `nutri_waybills.json`

### Step 2: 쇼핑몰주문번호 → 사방넷주문번호 매핑
- [ ] `classified.json`의 realship_orders에서 shmaOrdNo → ordNo 매핑
- [ ] 유니크 주문번호 확인 (한 주문에 여러 상품이면 송장 1개)

### Step 3: 사방넷 운송장 대량입력
- [ ] 엑셀 생성: [ordNo, wyblNo, "", "", 택배사코드]
- [ ] 사방넷 관리자 → 운송장입력(대량) → 택배사 선택 → 파일 업로드 → 저장
- [ ] 결과 팝업에서 "성공/실패" 건수 확인

### Step 4: 실패 건 재처리
- 실패 원인 "송장을 등록할 주문상태가 아닙니다" → 001 상태 남아있음
- [ ] 사방넷 주문서확정관리 → 미확정 건 확정 → 운송장입력 재시도
- 실패 주문번호만 따로 엑셀(`sabang_waybill_retry_N.xlsx`) 생성해 재업로드

### Step 5: 쇼핑몰 운송장 송신
- [ ] 쇼핑몰운송장송신 페이지
- [ ] 일자: WYBL_INPUT_DT / 오늘
- [ ] **"송장미송신"** 체크
- [ ] 검색
- [ ] Vue 테이블 체크박스 **toggleRowSelection**으로 전체선택
- [ ] **쿠팡 건**: 운송장송신 버튼 → 확인 모달 클릭 (자동 API)
- [ ] **스마트스토어 건**: 4단계 패턴 실행
- [ ] 각 송신 후 10초 대기 → 새로고침 → 재검색
- [ ] 송장미송신 건수 0건 될 때까지 반복

### Step 6: 종합 리포트 생성
- [ ] 빈박스(오전 처리분) + 실배송(오후 처리분) 전체 집계
- [ ] 풀필먼트 등록 / 송장배정 / 운송장입력 / 쇼핑몰송신 각 단계별 건수
- [ ] 특이사항 (매핑 실패 / 송신 실패 등) 상세

---

## 🔧 미리 체크해야 할 항목

### 풀필먼트 상품 매핑 (2026-04-24 아침 확인 필요)
- [ ] **바디인솔 아쿠아슈즈** — 풀필먼트에 SKU 등록됐는지?
  - 어제 주문 1건(2026042333953871, 블랙 260-265)이 매핑 실패
  - 등록됐다면 `product_mapping.json`에 키워드 규칙 추가
  - 미등록이면 사용자가 풀필먼트 상품 등록 먼저

### 가드웰 무릎보호대 옵션 확인
- [ ] 슬개골 무릎보호대 "수량: 한쪽" 주문이 오늘처럼 들어오면 **E00400639** (1개+마사지볼 L 블랙)로 자동 매핑됨 (어제 매핑 규칙 추가 완료)
- [ ] 다른 옵션(2p / 색상 별도) 들어오면 매핑 실패 가능 → 규칙 확장 필요

---

## 🆘 트러블슈팅 Quick Reference

| 문제 | 원인 | 해결 |
|------|------|------|
| 운송장입력 실패 "주문상태가 아닙니다" | 001 상태 | 주문서확정관리에서 일괄확정 |
| `updateOrdStsCd` API 500 | 사방넷 서버 이슈 | UI 자동화 또는 수동 |
| 쇼핑몰 송신 안 먹음 (스마트스토어) | 4단계 패턴 미사용 | 4단계 패턴 실행 |
| 4단계 패턴 두 번째 실행 실패 | iframe 캐시 | `window.__cleanIframe = null` + `location.reload()` |
| Vue 테이블 체크박스 클릭 무반응 | DOM만 바뀜, store 업데이트 안 됨 | `vm.toggleRowSelection(row, true)` |
| 풀필먼트 업로드 세션 만료 | 세션 짧음 | 계정별 재로그인 |
| 매핑 실패(unmapped) | 풀필먼트 SKU 미등록 or 규칙 누락 | SKU 확인 + product_mapping.json 규칙 추가 |

---

## 📁 관련 파일

- `outputs/run_tomorrow.py` — 주문 분류 실행기 (기본값: 어제 00:00 ~ 오늘 14:00)
- `outputs/convert_classified_to_sabang.py` — 한글헤더 → 영어키 변환
- `outputs/realship-fulfillment-plugin/skills/realship-fulfillment/scripts/create_fulfillment_excel.py` — 풀필먼트 엑셀 생성
- `outputs/realship-fulfillment-plugin/skills/realship-fulfillment/scripts/product_mapping.json` — 상품 매핑 규칙 (어제 덴코/글램루아/가드웰 추가됨)
- `outputs/realship-fulfillment-plugin/skills/realship-fulfillment/references/mapping_operation_rules.md` — 운영 원칙 (앞부분 일치 / 애매 카테고리 사전 확인 / 배타적 키워드)
