---
name: realship-fulfillment
description: 사방넷 → 풀필먼트 실배송 반자동화 스킬 (v0.5.6 — 전 매트릭스 lookup table화). 사용자가 사방넷에서 주문수집/일괄확정/엑셀 다운로드까지 직접 → 채팅에 엑셀 업로드 → Claude가 빈박스/실배송 분류 + 매핑 + 자가검증 보고서 → 사용자 OK 후 풀필먼트 업로드 엑셀(이더+뉴트리) 생성. 모든 룰 + 코드 lookup이 mapping_rules.json에 영구 박혀있어 재고조회 의존 없음. fixture_tests.json으로 매 사이클 회귀 테스트. 트리거: 실배송, 풀필먼트 등록, 발주등록, 풀필먼트 엑셀, 사방넷 발주, 스마트스토어 발주.
---

# 실배송 풀필먼트 (v0.5.7 — sku_names 영구화)

## v0.5.6 — 매트릭스 룰 전부 lookup table화 (2026-05-15)

재고조회 엑셀 의존을 **완전히 제거**. 사용자가 매 사이클 재고조회 다운로드 안 해도 됨.
- 구름깔창(루미솔/이더), 벌집깔창, 양말, 글램루아, 아쿠아슈즈 — 모든 matrix 룰의 코드를 `mapping_rules.json`의 lookup에 영구 박음 (총 85종)
- `resolve_matrix()` 일반화: 카테고리별 키 형식으로 lookup 우선 조회. mat 폴백은 신규 SKU 추가 대비용
- cross-validate에 **코드 변동 감지** 추가: 같은 주문번호인데 이전 사이클의 풀필먼트 코드와 다른 코드로 매핑되면 보고서에 알림 (정보용, stop X)
- **신규 SKU 추가 시**: mapping_rules.json의 해당 룰 lookup에 한 줄 추가

## v0.5.5 — 출고희망일 빈칸 + 우편번호 정규화 (2026-05-15)

- **출고희망일은 항상 빈칸**
- **우편번호 정규화 (`norm_zip`)**: 하이픈/공백 제거 + 5자리 zfill → leading 0 보존
- **우편번호 셀 text format** (`@`)

## 시스템 정보

### 사방넷 관리자 (사용자 직접)
URL: `https://sbadmin03.sabangnet.co.kr` / 로그인 `eithercompany / dlejzja7801!`

### 풀필먼트 (사용자 직접 업로드)
| 화주 | 회사코드 | 아이디 | 비번 |
|---|---|---|---|
| 이더컴퍼니 | w7298 | eithercompany | dlejrhddyd1@ |
| 뉴트리정 | w7298 | nutrijung | dlejrhddyd1@ |

URL: `https://wms02.sbfulfillment.co.kr`

## 워크플로우 (v0.5.6 간소화)

1. 사방넷 주문수집/일괄확정/엑셀 다운로드
2. 주문서 엑셀 1개만 채팅 업로드 (재고조회·발주조회 불필요)
3. Claude가 매핑 + cross-validate + 풀필먼트 엑셀 생성
4. 사용자 풀필먼트 사이트 업로드

```bash
python3 scripts/create_fulfillment_excel.py \
  --orders <주문서엑셀> \
  --prev-fulfillment-dir <Downloads또는워크스페이스폴더> \
  --state <Downloads>/realship_state.json --cycle-start "YYYY-MM-DD HH:MM" \
  --output-ether <이더엑셀> --output-nutri <뉴트리엑셀>
```

## 핵심 가드

1. shmaOrdNo 단일 키 매칭
2. 수량 룰 SKU N+N 셋트 단위
3. 빈박스 정확매칭
4. cross-validate 송장 중복 + 코드 변동 (v0.5.6)
5. 사이클 시작일 이전 SKIP
6. 헤더 23 컬럼 + 출고희망일 빈칸 (v0.5.5)
6b. 우편번호 정규화 (norm_zip + text format)
7. 이더(E*) vs 뉴트리(N*) 분리
8. 사용자 fallback (010-0000-0000 / 00000)
9. --report-only 모드 confirm 게이트
10. auto_stop_thresholds 자동 stop
11. fixture 회귀 테스트
12. 룰 충돌 감지
13. 합배송 정렬 (받는분+전화)
14. 코드 변동 감지 (v0.5.6)

## 변경 이력

- **v0.5.7** (2026-05-18) — **sku_names 영구 박음** (108종, E*+N*). 재고조회 엑셀 없어도 풀필먼트 엑셀 판매상품명이 정상 표시. to_xlsx 폴백 우선순위: SKU master > mapping_rules.sku_names > 룰 id.
- **v0.5.6** (2026-05-15) — 전 매트릭스 룰 lookup 영구화 (85종). resolve_matrix 일반화. 코드 변동 감지.
- **v0.5.5** (2026-05-15) — 출고희망일 빈칸 + 우편번호 norm_zip + text format. 아쿠아슈즈 lookup 28종.
- **v0.5.4** (2026-05-15) — 아쿠아슈즈 lookup table화.
- **v0.5.3** (2026-05-13) — 어제 풀필먼트 엑셀 자동 cross-validate.
- **v0.5.0** (2026-05-08) — 합배송 정렬.
- **v0.4.0** (2026-05-07) — 데이터 드리븐 룰 엔진.
