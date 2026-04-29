---
name: plugin-github-sync
description: >
  창근님 `realship-fulfillment-plugin` 플러그인을 수정했을 때 GitHub 레포(EitherCompany/realship-fulfillment-plugin)로 자동 푸시 + 릴리스 태그까지 수행한다.
  PAT 는 노션 "👮 이창근" 페이지에서 자동 조회 (플러그인에 토큰 하드코딩 절대 금지).
  반드시 이 스킬을 사용해야 하는 경우: "실배송 플러그인 수정", "실배송 플러그인 업데이트", "스킬 수정해줘", "SKILL.md 고쳐", "재패키징", "배포",
  "버전 올려", "마켓플레이스 업데이트", "GitHub 푸시", "플러그인 재배포", "v0.x.x 올려줘" 등 플러그인 편집 후 GitHub·마켓플레이스 반영과 관련된 모든 요청.
---

# realship-fulfillment-plugin GitHub 자동 동기화 (plugin-github-sync · v1.0.0)

## 자동 트리거

플러그인 내 다음 중 하나라도 변경되면 별도 지시 없이 절차 착수:
- `.claude-plugin/plugin.json` · `marketplace.json` 편집
- `skills/*/SKILL.md` 편집 또는 신규 스킬 추가
- `skills/*/references/*` · `skills/*/scripts/*` 편집
- `README.md` 편집

## 고정 상수

- **레포**: `EitherCompany/realship-fulfillment-plugin` (Private)
- **PAT**: 노션 페이지 `1d8d9e75-0367-80b3-9f32-e82210a58e20` 에서 `github_pat_...` 정규표현식 추출
- **기본 브랜치**: `main`

## 레포 구조 (binbox 패턴)

```
realship-fulfillment-plugin/                   ← GitHub 레포 루트
├── .claude-plugin/
│   └── marketplace.json                        ← 루트 marketplace ($schema + owner + plugins[])
├── README.md
└── realship-fulfillment-plugin/                ← 실제 플러그인 폴더
    ├── .claude-plugin/
    │   └── plugin.json                         ← author 객체 + license:MIT + keywords
    └── skills/
        ├── realship-fulfillment/
        │   ├── SKILL.md
        │   ├── state.json
        │   ├── references/
        │   └── scripts/
        └── plugin-github-sync/
            └── SKILL.md
```

**3 곳 version 동기 bump 필수**:
- 루트 `.claude-plugin/marketplace.json` 의 최상위 `version`
- 같은 파일의 `plugins[0].version`
- `realship-fulfillment-plugin/.claude-plugin/plugin.json` 의 `version`

## 자동 실행 절차

### Step 1 — PAT 조회

`notion-fetch(id: "1d8d9e75-0367-80b3-9f32-e82210a58e20")` → 본문 정규표현식 `github_pat_...` 추출.

### Step 2 — 변경 확인 + 버전 결정 (semver)

- patch (x.y.N): 문구·버그 픽스
- minor (x.Y.0): 새 정책·스킬·기능
- major (X.0.0): 구조 변경

### Step 3 — 3 곳 version bump

### Step 4 — git clone (임시 디렉토리)

```bash
TMP="/tmp/.realship_push_$$"
git clone --depth 1 "https://x-access-token:${TOKEN}@github.com/EitherCompany/realship-fulfillment-plugin.git" "$TMP"
```

### Step 5 — 변경본 덮어쓰기

```bash
rsync -av --delete --exclude='.git' --exclude='__pycache__' "$SRC/" "$TMP/"
```

### Step 6 — Secret 검사

```bash
HITS=$(grep -rEn 'github_pat_[A-Za-z0-9_]{50,}' "$TMP" --exclude-dir=.git)
if [ -n "$HITS" ]; then echo "🚨 leak"; exit 1; fi
```

### Step 7 — 커밋·태그·푸시

```bash
cd "$TMP"
git config user.email "ghgh404@gmail.com"
git config user.name "이창근"
git add -A
git commit -m "v<N>.<M>.<K>: <한 줄 요약>"
git push origin main
git tag -a v<N>.<M>.<K> -m "<요약>"
git push origin v<N>.<M>.<K>
```

### Step 8 — GitHub 릴리스

`POST /repos/EitherCompany/realship-fulfillment-plugin/releases` (Bearer ${TOKEN}).

### Step 9 — 정리

```bash
rm -rf "$TMP"   # PAT 잔존 방지
```

## 금지 사항

1. GitHub 웹 UI 자동화 편집 금지
2. PAT 를 플러그인·커밋·로그에 노출 금지
3. 한 version 필드만 bump 금지 — 3 곳 모두 동기 bump
4. `.git` 폴더 잔존 push 금지

## 완료 기준

- [ ] 노션 PAT 조회 완료
- [ ] 3 곳 version 동기 bump
- [ ] Secret 검사 통과
- [ ] `git push origin main` + `git push origin v<N>.<M>.<K>` 성공
- [ ] GitHub 릴리스 생성 성공
- [ ] 임시 디렉토리 정리
