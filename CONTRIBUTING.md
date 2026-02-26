# 🤝 기여 가이드 (Contributing Guide)

Total-HR-System 프로젝트에 기여해 주셔서 감사합니다!

---

## 🌿 브랜치 전략 (Branch Strategy)

이 프로젝트는 **Git Flow** 방식을 따릅니다.

```
main          ← 배포(프로덕션) 브랜치. 직접 커밋 금지.
develop       ← 통합 개발 브랜치. 모든 기능은 여기서 합쳐집니다.
feature/*     ← 기능 개발 브랜치 (예: feature/staff-search)
hotfix/*      ← 긴급 버그 수정 브랜치 (예: hotfix/login-bug)
```

---

## 🚀 개발 워크플로우

### 1. 새 기능 개발
```bash
# develop 브랜치에서 feature 브랜치 생성
git checkout develop
git pull origin develop
git checkout -b feature/기능명

# 작업 완료 후 develop으로 PR 생성
```

### 2. 긴급 버그 수정
```bash
# main 브랜치에서 hotfix 브랜치 생성
git checkout main
git checkout -b hotfix/버그명

# 수정 완료 후 main과 develop 모두 PR 생성
```

### 3. 배포 (develop → main)
- develop 브랜치에서 main으로 PR을 생성하고 코드 리뷰 후 병합합니다.

---

## 📝 커밋 메시지 규칙

```
[타입] 변경 내용 요약

예시:
[feat] 스텝 계약 갱신 알림 기능 추가
[fix] 성과평가 저장 오류 수정
[refactor] Code.gs 함수 분리 리팩토링
[docs] CONTRIBUTING.md 업데이트
[style] 인덴트 및 공백 정리
```

| 타입 | 설명 |
|------|------|
| feat | 새로운 기능 추가 |
| fix | 버그 수정 |
| refactor | 코드 리팩토링 |
| docs | 문서 변경 |
| style | 코드 포맷팅 (기능 변경 없음) |
| test | 테스트 추가/수정 |
| chore | 빌드, 설정 변경 |

---

## 🔍 Pull Request 규칙

1. PR은 반드시 `develop` 브랜치를 대상으로 생성합니다 (hotfix 제외)
2. PR 제목은 커밋 메시지 규칙을 따릅니다
3. PR 템플릿을 빠짐없이 작성합니다
4. 최소 1명의 리뷰어 승인 후 병합합니다
5. 병합 후 feature 브랜치는 삭제합니다

---

## 📁 프로젝트 구조

```
Total-HR-System/
└── 지수_스텝관리/          # 지수회계법인 환급사업부 스텝관리 시스템
    ├── Code.gs             # 백엔드 메인 로직
    ├── Index.html          # 프론트엔드 UI
    ├── PatchStaffContacts.gs  # 연락처 일괄 업데이트
    └── SeedStaffData.gs    # 초기 데이터 시드
```
