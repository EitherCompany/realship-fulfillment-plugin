# realship-fulfillment-plugin (v0.2.4)

이더컴퍼니 실배송 풀필먼트 자동 등록 플러그인 — **사방넷 → 풀필먼트 등록까지만**.

> v0.2.4 변경: 송장처리/쇼핑몰송신 단계 제거 (별도 스킬로 분리). 풀필먼트 등록까지만 안전하게 자동화.

## 포함 스킬

- **realship-fulfillment** — 매일 1차/2차 사이클로 사방넷 주문수집 → 일괄확정 → 001→002 → 분류·매핑 → 풀필먼트(이더+뉴트리) 발주등록 엑셀 생성 + 사용자 업로드.
- **plugin-github-sync** — 플러그인 편집 시 GitHub 자동 푸시 + 릴리스 태그.

## 핵심 룰 (v0.2.4)

1. **shmaOrdNo 단일 키 매칭** — 받는분 이름 매칭 절대 금지 (cross-account 사고 차단)
2. **수량 룰 v2** — `풀필 수량 = ordQty × set_multiplier(옵션)` (1+1/2+2/3+3/N개)
3. **빈박스 룰 정정 (v0.2.4)** — `'문 앞에 놓아주세요!' in msg` 느낌표 정확 매칭만 빈박스. 부분 매칭 금지!
   - 04-29 사고: 부분 매칭으로 9건(스마트스토어 일반 안전 메시지) 누락 직전 → 스마트스토어 크로스체킹으로 캐치
4. **state.json 차집합 (v0.2.4)** — `same shmaOrdNo + same 사방넷코드` 모두 매칭 시에만 SKIP. 옵션 다른 추가주문 가드.
5. **04-28 같은 사이클 시작일 이전 주문 자동 SKIP** (사용자 수동처리분 중복 차단)
6. **헤더 강제 (v0.2.4)** — 풀필먼트 발주 엑셀은 스크립트의 HEADERS 상수 import 사용. 직접 작성 금지.
7. **빈 전화/우편/주소 fallback** — 010-0000-0000 / 00000 / "주소 사방넷 자동입력"
8. **출고희망일 자동 = 다음날 (YYYY-MM-DD)**
9. **스마트스토어 일주일치 크로스체킹 의무 (v0.2.4 신규)** — 풀필먼트 업로드 전 누락/중복 캐치

## 시스템 정보

| 항목 | 값 |
|---|---|
| 사방넷 URL | https://sbadmin03.sabangnet.co.kr |
| 사방넷 로그인 | eithercompany / dlejzja7801! |
| 풀필먼트 URL | https://wms02.sbfulfillment.co.kr |
| 풀필먼트 회사코드 | w7298 |
| 풀필먼트 비밀번호 | dlejrhddyd1! (이더/뉴트리 통일) |

## 자동화 한계 (Chrome MCP 환경)

| 항목 | 자동화 | 사용자 작업 |
|---|---|---|
| 사방넷 주문수집 | ✅ 자동 | - |
| 사방넷 일괄확정 | ✅ 자동 | - |
| 사방넷 001→002 변환 | ✅ 자동 | - |
| 주문서확인처리 엑셀 다운 | ❌ | 30초 (양식1 + 전체자료) |
| 분류·매핑·엑셀 생성 | ✅ 자동 | - |
| 풀필먼트 사이트 업로드 | ❌ | 5분 (이더 → 뉴트리 두 계정) |

## 설치 / 업데이트

```
/plugin install realship-fulfillment-plugin@EitherCompany
/plugin update realship-fulfillment-plugin@EitherCompany
```

## 변경 이력

[GitHub Releases](https://github.com/EitherCompany/realship-fulfillment-plugin/releases)

