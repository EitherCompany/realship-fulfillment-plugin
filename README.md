# realship-fulfillment-plugin (v0.4.0)

**데이터 드리븐 룰 엔진 + 충돌 자동 감지 + fixture 회귀 테스트**

## v0.4.0 메이저 변경

### 데이터 드리븐 룰 엔진
모든 매핑 룰이 `scripts/mapping_rules.json` 한 파일에. **Python 코드 수정 없이 새 상품 추가 가능.**

```json
{
  "mapping_rules": [
    {"id":"베개_덴코",   "priority":20, "all":["덴코","베개"],          "code":"E00400497"},
    {"id":"베개_경추",   "priority":22, "all":["경추","베개"], "none":["덴코"], "code":"E00400015"},
    ...
  ],
  "quantity_rules": [...],
  "binbox_rules": [...],
  "auto_stop_thresholds": {...}
}
```

### 단일 평가 엔진 + 충돌 감지
- 우선순위 정렬 후 첫 매칭 룰 사용
- 동일 우선순위 다중 매칭 시 **충돌 감지** → 자동 stop
- if/elif 분기 제거 → 룰 추가가 안전

### fixture 회귀 테스트
`scripts/fixture_tests.json` 20개 매핑 + 4개 주소 + 5개 빈박스. 매 사이클 시작 시 자동 통과 의무.

### 잔재 4개 삭제
- `references/mapping_operation_rules.md` (v0.2.x)
- `references/tomorrow_flow_checklist.md` (2026-04-24)
- `scripts/match_waybills.py` (송장처리 별도 스킬)
- `state.json` (사용자 워크스페이스에 두라고 SKILL에 명시)

## 새 상품 추가하는 법

mapping_rules.json에 룰 객체 한 줄:

```json
{"id":"새상품_X","priority":35,"all":["키워드1","키워드2"],"code":"E00400999"}
```

fixture_tests.json에 검증 케이스 추가:

```json
{"id":"새상품_X_test","product":"...","option":"...","expected_code":"E00400999","expected_qty_per_ord":1}
```

→ 코드 수정 0줄. 다음 사이클부터 적용.

## 워크플로우 (자세한 절차)

`realship-fulfillment-plugin/skills/realship-fulfillment/SKILL.md` 참조.

## 변경 이력

- **v0.4.0** (2026-05-07) — 데이터 드리븐 룰 엔진, fixture 20개, 잔재 삭제
- **v0.3.1** (2026-05-07) — 9가지 결함 fix
- **v0.3.0** (2026-05-06) — 메이저 반자동화 전환
- v0.2.5 (2026-05-04) / v0.2.4 (2026-04-29) / v0.2.0 (2026-04-28)
