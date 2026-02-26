# 🌿 GESI Annual Report Generator

**녹색에너지전략연구소(GESI)** 의 연례보고서를 자동으로 생성하는 Python 스크립트입니다.  
엑셀 데이터베이스(`Annual Report Database.xlsx`)를 읽어 Word 문서(`.docx`) 형태의 연례보고서를 만들어냅니다.

---

## 📁 프로젝트 구조

```
Annual_Report/
├── report_gen.py                   # 메인 보고서 생성 스크립트
├── Annual Report Database.xlsx     # 연구 데이터베이스 (입력 파일)
├── toc_background.png              # 목차 페이지 배경 이미지 (AI 생성)
└── README.md
```

### 엑셀 데이터베이스 시트 구성

| 시트명 | 설명 |
|---|---|
| `Institute_Info` | 기관 소개 (연도별 운영 개요, 인력, 연락처) |
| `Research_History` | 연도별 연구 연혁 (3개 Pillar 구분) |
| `Master_Research` | 개별 연구 상세 (ID, 분야, 제목, 요약, 발주처 등) |
| `Activities_&_News` | 주요 활동·행사·외부 소통 이력 |
| `2025` | 해당 연도 Pillar별 상세 서술 (소제목, 연구배경, 요약, 효과) |

---

## 🚀 실행 방법

### 1. 의존성 설치

```bash
pip install pandas python-docx openpyxl matplotlib
```

### 2. 실행

```bash
python report_gen.py
```

실행 후 `GESI_2025_Annual_Report_Final.docx` 파일이 생성됩니다.

---

## 📄 보고서 구성

| 섹션 | 설명 |
|---|---|
| **표지** | 연도, 슬로건 |
| **목차** | 전체 목차 + AI 생성 비전 이미지 |
| **01. 기관 소개** | `Institute_Info` 시트 기반 |
| **02. 연구 연혁** | `Research_History` 시트 + 타임라인 인포그래픽 |
| **03. 2025 핵심 연구 성과** | `Master_Research` 시트 (Pillar별 분류) |
| **04. 주요 활동 및 외부 소통** | `Activities_&_News` 시트 |
| **05. 2025년 주요 연구 상세** | `2025` 시트 (Pillar당 1페이지) |

---

## 🔑 주요 기능

- **엑셀 → Word 자동 변환**: 시트별 데이터를 섹션별로 자동 매핑
- **타임라인 인포그래픽**: matplotlib으로 2015–2025 연구 흐름 시각화
  - 3개 Pillar 레인 (시스템 전환, 지역 에너지 전환, 탈탄소 정책)
  - 버블 크기 = 연도·분야별 연구 건수
  - `연구 분야 ID` 기반 연구 연계/확장 관계 화살표 표시
- **Pillar별 상세 페이지**: `2025` 시트의 각 Pillar 컬럼을 1페이지씩 구성
- **연도 자동 확장**: 새 연도 시트(`2026` 등)를 추가하면 자동 반영 가능

---

## 🛠 클래스 구조

```
GesiFullReportGenerator
├── add_title_page()               # 표지
├── add_toc_page(image_path)       # 목차 (이미지 포함)
├── add_institute_intro()          # 기관 소개
├── add_research_history()         # 연구 연혁 + 인포그래픽
│   └── _create_timeline_infographic()  # matplotlib 인포그래픽 생성
├── add_key_research_2025()        # 핵심 연구 성과
├── add_activities()               # 주요 활동
├── add_2025_pillar_pages()        # Pillar별 상세 페이지
└── save_report(filename)          # Word 저장
```

---

## 📦 의존 패키지

| 패키지 | 용도 |
|---|---|
| `pandas` | 엑셀 데이터 읽기 |
| `openpyxl` | `.xlsx` 파일 파싱 엔진 |
| `python-docx` | Word 문서 생성 |
| `matplotlib` | 타임라인 인포그래픽 생성 |

---

## 📝 커스터마이징

새 연도 보고서를 만들려면:
1. `Annual Report Database.xlsx`에 해당 연도 시트(`2026` 등) 추가
2. `report_gen.py` 상단의 연도 문자열(`2025`) 수정
3. `add_2025_pillar_pages()` 메서드의 시트명 수정

---

> **GESI (Green Energy Strategy Institute)**  
> 02-552-0940 / gesi@gesi.kr
