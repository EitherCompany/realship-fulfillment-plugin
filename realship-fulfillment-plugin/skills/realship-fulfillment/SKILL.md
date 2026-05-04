---
name: realship-fulfillment
description: >
  실배송 풀필먼트 자동 등록 스킬 (v0.2.5 — 풀필먼트 업로드 자동화 추가 / 사용자 작업: 사방넷 다운로드 30초만).
  매일 1차/2차 사이클로 사방넷 → 사방넷 풀필먼트(이더+뉴트리)에 실배송 자동 등록.
  반드시 이 스킬을 사용해야 하는 경우: "실배송 풀필먼트", "풀필먼트 등록", "실배송 등록", "오늘 실배송",
  "풀필먼트 발주", "실배송 처리해줘", "주문 마감", "2시 마감", "풀필먼트에 넣어줘", "실배송 올려줘",
  "발주등록해줘", "실배송 1차", "실배송 2차", "사방넷 발주", "스마트스토어 발주", "발주 엑셀 만들어줘"
---

# 실배송 풀필먼트 자동 등록 (v0.2.5)

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

⚠️ **풀필먼트 발주조회 엑셀 비밀번호 보호** — 위 비번으로 msoffcrypto-tool 복호화

---

## v0.2.5 절대 룰 (사고 기반 영구 반영)

1. **shmaOrdNo 단일 키 매칭** — 받는분 이름 매칭 절대 금지 (김채은 cross-account 사고)
2. **수량 룰 v2** — `풀필 수량 = ordQty × set_multiplier(옵션)` (1+1/2+2/3+3/N개)
3. **빈박스 룰** — `'문 앞에 놓아주세요!' in msg` 느낌표 정확 매칭만 빈박스 (부분 매칭 금지)
4. **사전 일괄확정·001→002 자동** (Vue VM 트릭)
5. **풀필먼트 엑셀 생성 후 사용자 수량 검토 의무** (미분류 발생 시 SKU 마스터 자동 활용)
6. **사이클 시작일 이전 주문 자동 SKIP** (수동처리분 중복 차단)
7. **state.json 차집합** — `shmaOrdNo + 사방넷코드` 둘 다 매칭 시에만 SKIP (옵션 다른 추가주문 가드)
8. **풀필먼트 발주 엑셀 헤더 강제** — 스크립트 HEADERS 상수 import (직접 작성 금지)
9. **빈 전화/우편/주소 fallback** — 010-0000-0000 / 00000 / "주소 사방넷 자동입력"
10. **출고희망일 자동 = 다음날** (YYYY-MM-DD)
11. **수집기간 6일치 강제** — 사방넷 자동수집은 빨간날·주말 동안 멈춤. 사이클 시작 시 사용자 명시 시작일 + 마진 2일 = 6일치 수동 트리거
12. **풀필먼트 SKU 마스터 자동 활용** — 미분류 발생 시 사용자 재고조회 다운 → 자동 매핑 → product_mapping.json 갱신

---

## 자동화 매트릭스 (v0.2.5)

| Step | 자동화 | 사용자 작업 | 시간 |
|---|---|---|---|
| 1. 주문수집 | ✅ | - | 0초 |
| 2. 일괄확정 | ✅ | - | 0초 |
| 3. 001→002 | ✅ | - | 0초 |
| 4. 사방넷 다운로드 | ❌ | 양식1 + 전체자료 다운 + 채팅 업로드 | **30초** |
| 5. 분류·매핑·차집합 | ✅ | - | 10초 |
| 5+. 미분류 SKU 마스터 (필요 시) | ⚠️ | 재고조회 엑셀 다운 (일주일 1회) | 30초 |
| 6. 풀필먼트 업로드 (이더+뉴트리) | ✅ | - | 0초 |
| 7. state.json 갱신 | ✅ | - | 0초 |

**총 사용자 작업: 30초 ~ 1분**

---

## 매 사이클 표준 7단계

### Step 1 — 사방넷 주문수집 (자동)

```javascript
// vm 검색
let cvm=null;
function walk(c,d=0){if(!c||d>50||cvm)return;if(c.$data && Array.isArray(c.$data.tableData) && c.$data.tableData.length===7 && c.$data.tableData[0]?.shmaId){cvm=c;return;}if(c.$children)c.$children.forEach(x=>walk(x,d+1));}
walk(document.getElementById('app').__vue__);

// 수집기간 6일치 강제 (빨간날·주말 누락 대응)
cvm.setDate(6);  // 또는 sbForm.startDate/endDate 직접 set
cvm.sbForm.startDate = '20260429';  // 사이클 시작 - 마진 1일
cvm.sbForm.endDate = '20260504';

// 7개 mall toggleRowSelection
const tables = document.querySelectorAll('.el-table');
let target = null;
tables.forEach(t => { const v = t.__vue__; if (v && v.store && v.store.states.data.length === 7) target = v; });
target.store.states.data.forEach(r => target.toggleRowSelection(r, true));

// 트리거
cvm.popOpenOrderCollect();
// 확인 모달 (880, 494 또는 880, 517)
// 5분 대기 (빨간날 누적 따라잡는 시간)
```

### Step 2 — 사방넷 일괄확정 (자동)

```javascript
let dvm=null;
function walk(c,d=0){if(!c||d>60||dvm)return;if(c.$data && Object.keys(c.$data).includes('sbForm') && Object.keys(c.$data).includes('tableData') && Object.keys(c.$data).includes('checkList')){dvm=c;return;}if(c.$children)c.$children.forEach(x=>walk(x,d+1));}
walk(document.getElementById('app').__vue__);

dvm.getOrderDecideSearch();  // 검색
// 검증: dvm.sbForm.total 0건이면 모두 002 (정상)
// > 0이면 일괄주문확정 (1308, 290) → 모달 (768, 410) → 확인 (880, 386)
```

### Step 3 — 001→002 변환 (자동, Vue VM 트릭)

기존 popup VM 트릭 유지 (binbox 패턴). 실제로 사방넷 자동봇이 처리하므로 신규주문 0건이면 SKIP.

### Step 4 — 주문서확인처리 엑셀 다운로드 (사용자 직접)

**한계**: `vm.makeExcelDownload()` 호출 시 `searchForm: {}` 빈 객체로 전송 → 서버 측 `code 10000` 거부. vm 인스턴스 sbForm reset 문제. Chrome MCP 환경에서 우회 불가능.

**워크어라운드** (30초):
1. 주문서확인처리 → 송장미등록 체크 → 검색
2. 양식 dropdown → 양식1 / 전체자료
3. 다운로드 → 채팅 업로드

### Step 5 — 분류·매핑·차집합 (자동)

```bash
python3 scripts/create_fulfillment_excel.py \
  --orders today_orders.json \
  --mapping scripts/product_mapping.json \
  --state <workspace>/realship_state.json \
  --output fulfillment_$(date +%Y%m%d).xlsx
```

자동:
- 빈박스 SKIP (쿠팡 % / 스마트스토어 + `'문 앞에 놓아주세요!'` 정확)
- 사이클 시작일 이전 SKIP
- state.json 차집합 (어제 등록 shmaOrdNo + 코드 둘 다 매칭)
- 매핑: 글램루아 24코드 매트릭스 + 아쿠아슈즈 28코드 매트릭스 + 깔창/뉴트리 키워드

### Step 5+ — 미분류 SKU 마스터 자동 매핑 (필요 시)

미분류 발생 시:
1. 사용자에게 재고조회 엑셀 다운 부탁 (일주일 1회 정도)
2. 스크립트가 자동 파싱 → 글램루아/아쿠아슈즈/뉴트리 SKU 검색
3. 매칭된 코드 자동 적용 + product_mapping.json 갱신

### Step 6 — 풀필먼트 업로드 (자동, v0.2.5 신규)

```javascript
// 1. 풀필먼트 로그인 (토큰 우선)
// 2. /order/add 직접 진입 (대시보드 거치면 hydration OK)
// 3. 엑셀등록 모달 → 파일 업로드 → 등록
// 4. 회사 dropdown 변경 (이더 → 뉴트리) 후 반복

// 검증: "엑셀 업로드 최근 이력"에 파일명 + 시간 + 크기 매칭
```

**핵심 발견 (v0.2.5)**: 풀필먼트 사이트는 토큰 살아있을 때 + 대시보드 거쳐 진입하면 Vue 3 hydration 정상. 이전 사이클 spinner 막힘은 콜드 스타트 + 토큰 만료 복합 원인.

### Step 7 — state.json 갱신 + 종합 보고서 (자동)

```json
{
  "last_cycle_date": "2026-05-04",
  "2026-05-04": {
    "ether_pairs": [["shma1", "code1"], ...],
    "nutri_pairs": [...],
    "all_pairs": [...]
  }
}
```

state 위치: 사용자 워크스페이스 (`<Downloads>/realship_state.json`). 플러그인 폴더에는 저장 금지.

---

## 핵심 가드 (CRITICAL)

1. **shmaOrdNo 단일 키 매칭** (받는분 이름 매칭 절대 금지)
2. **수량 룰 v2** — set_multiplier(옵션) 적용
3. **빈박스 룰** — `'문 앞에 놓아주세요!' in msg` 느낌표 정확 매칭만
4. **state.json 차집합** — shmaOrdNo + 코드 둘 다 매칭 시에만 SKIP
5. **6일치 강제 수집** — 사방넷 자동수집은 빨간날·주말 동안 멈춤
6. **풀필먼트 업로드 자동** — 토큰 살아있을 때 대시보드 거쳐 진입

---

## 변경 이력

- **v0.2.5** (2026-05-04) — **풀필먼트 업로드 자동화 추가**, 6일치 강제 수집 룰, 글램루아 24코드 + 아쿠아슈즈 28코드 매트릭스, SKU 마스터 자동 매핑, 비번 정정 dlejrhddyd1@, 빨간날·주말 자동수집 누락 대응.
- **v0.2.4** (2026-04-29) — 송장처리 제거, 빈박스 룰 정정 (`!` 정확), 매핑 7종 추가, state.json 차집합, 스마트스토어 크로스체킹.
- **v0.2.3** (2026-04-28) — 폴더 구조 정정, plugin/marketplace 스키마 정정.
- **v0.2.2** (2026-04-28) — plugin-github-sync 스킬 추가.
- **v0.2.1** (2026-04-28) — 중복 발주 가드.
- **v0.2.0** (2026-04-28) — 수량 룰 v2, shmaOrdNo 단일 키, 사전 일괄확정 의무화.

