# 원광대학교 한의과대학 경혈학 학습일지 작성 가이드 및 다면평가 시스템

> **경혈학 학습일지 온라인 가이드 웹서비스 및 자동화 다면평가 프레임워크 (v3.2)**  
> 학생 안내용 GitHub Pages 웹 포털 & AI 에이전트(Antigravity·Claude Code) 평가 파이프라인

[![GitHub Pages](https://img.shields.io/badge/GitHub_Pages-Online-emerald?style=flat&logo=github)](https://jaehyokim-sketch.github.io/Study-Log/)
[![Version](https://img.shields.io/badge/Rubric-2026.10.23_Revised-blue?style=flat)](references/rubric.md)
[![Python 3.9+](https://img.shields.io/badge/Python-3.9+-green.svg)](scripts/eval_engine.py)

---

## 🌐 온라인 작성 가이드 (GitHub Pages)

학생들은 아래 웹페이지에서 최신 평가기준과 클로드 베르나르 5단계 분석법이 반영된 공식 가이드를 실시간으로 열람할 수 있습니다:
👉 **[경혈학 학습일지 작성 가이드 온라인 포털](https://jaehyokim-sketch.github.io/Study-Log/)**

---

## 📌 주요 특징 및 평가 체계

### 1. 학생용 작성 가이드 핵심 원칙
- **4대 핵심 섹션 구조**:
  1. 수업 내용 요약 (30%) - 동영상 강의를 자신의 언어로 재구성
  2. 자가 학습/탐구 학습 (40%) - **클로드 베르나르의 5단계 과학적 방법론**(질문·조작·측정·결과·해석) 적용
  3. 요약 및 토론 (30%) - 종합 정리, 비판적 의견 및 심화 의문
  4. 출처 및 참고문헌·AI 활용 표기 (필수)
- **AI 활용 투명성 규범 (A1~A8)**:
  - 도구명(구체적 모델 버전), 활용 목적, 입력 프롬프트, AI 결과물 대비 본인 수정 내역, 적용 위치(쪽수) 상세 기재
  - 원광대학교 중앙도서관 CopyKiller 연계 검증 안내
- **2026-10-23 완화 개정 루브릭 및 형성평가 반영**:
  - 5단계 중 4/5 충족자라도 기본 요건(10쪽 이상 분량, 충실한 성찰, 출처) 충족 시 '우수' 유연 수용
  - 학과 배포 **표준 템플릿(소제목 구조)** 준수 권고
  - 문헌 고찰보다는 **실험 연구 또는 임상 연구(RCT 등)** 논문 우선 선정 권장

### 2. 교수자/조교용 자동화 평가 시스템 (v3.2)
- **4단계 관문(Gate) 판정 절차**:
  `미제출 확인 → 8대 결격 사유(D1~D8) 확인 → 6대 우수 요건(R1~R6) 확인 → 매우 우수(교수자 정성평가) → 상한 적용`
- **다중 시트 엑셀 리포트 자동 산출**:
  1. `종합요약`: 등급 분포 통계 및 시각화용 요약
  2. `상세평가`: 학생별 23개 지표별 세부 채점표
  3. `교수자피드백`: 학생별 강점/보완점/맞춤형 가이드라인

---

## 🗂️ 폴더 구조

```text
C:\workspace\학습일지 평가\ (Repository: Study-Log)
├── index.html                # 학생 안내용 온라인 가이드 웹페이지 (GitHub Pages)
├── 학습일지_평가지침.md        # 공식 평가지침 및 판정 기준 전문
├── README.md                 # 본 프로젝트 개요 및 운영 가이드
├── 설치_및_설정_안내.md         # 상세 설치, 환경설정 및 트러블슈팅 매뉴얼
├── install.ps1               # 원클릭 AI 스킬 자동 배포 스크립트
├── deploy_to_gdrive.py       # Google Drive 양방향 동기화 배포 스크립트
│
├── references\
│   └── rubric.md             # 2026-10-23 완화 개정판 세부 판정 루브릭
│
└── scripts\
    ├── eval_engine.py        # 핵심 평가 엔진 및 엑셀 리포트 자동 생성기
    └── rename_logs.py        # 파일명 일괄 표준화 (회차_이름_학번_경맥명 또는 학습주제.pdf)
```

---

## 🚀 빠른 시작 (Quick Start)

### 1. 스킬 자동 설치 (Antigravity & Claude Code)
```powershell
cd "C:\workspace\학습일지 평가"
.\install.ps1
```

### 2. CLI 명령어 사용법
```bash
# 파일명 일괄 정리
python scripts/rename_logs.py "대상폴더" --round 6 --meridian "독맥"

# 전수 평가 및 엑셀 리포트 생성
python scripts/eval_engine.py --target "대상폴더" --prev4 "4차결과.xlsx" --prev5 "5차결과.xlsx" --round 6 --meridian "독맥" --output "최종피드백.xlsx"
```

---

## 💬 대화형 AI 어시스턴트 호출

대화창에서 아래 키워드로 평가 스킬을 활성화할 수 있습니다:
- `[학습일지평가]`
- `[학습일지 평가]`
- `@학습일지평가`

---

**작성 및 관리:** 원광대학교 한의과대학 경혈학교실  
**최종 갱신:** 2026-11-04 (2026-10-23 개정 루브릭 반영)
