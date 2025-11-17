# Word 산출물 자동 생성 프로그램

YAML 파일로부터 전문적인 한국어 기술 문서를 자동으로 Word 문서(.docx)로 생성하는 프로그램입니다.

<img width="675" height="193" alt="1" src="https://github.com/user-attachments/assets/16078f0f-2772-4613-beb4-b98fde820bac" />
<img width="626" height="795" alt="2" src="https://github.com/user-attachments/assets/75e7fbf5-10a0-4abe-b224-c0840c6df48b" />
<img width="435" height="525" alt="3" src="https://github.com/user-attachments/assets/ec6d39ec-3b9c-47e2-bc95-a9c44492b72c" />

## 주요 기능

- ✅ **표지 페이지**: 문서 제목과 프로젝트 정보 표 자동 생성
- ✅ **하이퍼링크 목차**: 점선 연결, 클릭 가능한 목차 (밑줄 없음)
- ✅ **페이지 번호**: 푸터에 "현재페이지 / 전체페이지" 자동 삽입
- ✅ **자동 필드 업데이트**: 문서를 열면 페이지 번호 자동 업데이트
- ✅ **계층적 구조**: 재귀적 섹션 구조 지원 (무제한 깊이)
- ✅ **새 페이지 구분**: Heading 1마다 자동으로 새 페이지 시작

## 생성되는 문서 구조

```
📄 1페이지: 표지
  - 문서 제목 (28pt, 중앙 정렬)
  - 프로젝트 정보 표 (프로젝트명, 고객사, 작성일자, 작성자, 버전)

📋 2-4페이지: 목차
  - 점선으로 연결된 목차 (제목 ........ 페이지번호)
  - 클릭 가능한 하이퍼링크 (밑줄 없음)
  - 계층적 들여쓰기

📝 5페이지 이후: 본문
  - Heading 1마다 새 페이지 시작
  - 계층적 제목 구조 (Heading 1, 2, 3, 4, 5)
  - 단계별 설명 (List Paragraph 스타일)
  - 일반 텍스트 내용 (content 필드)

📄 푸터: 페이지 번호
  - 중앙 정렬, "페이지번호 / 총페이지수" 형식
```

## 설치 및 실행

### 1. 필수 요구사항

- Python 3.8 이상
- pip (Python 패키지 관리자)

### 2. 라이브러리 설치

```bash
pip install -r requirements.txt
```

필요한 라이브러리:
- `python-docx==1.1.0`: Word 문서 생성
- `PyYAML==6.0.1`: YAML 파일 파싱

### 3. 프로그램 실행

```bash
python create_document.py
```

실행하면 `output.docx` 파일이 생성됩니다.

## 데이터 형식 (YAML)

`sample_data.yaml` 파일을 수정하여 원하는 내용을 작성합니다.

### YAML 구조

```yaml
metadata:
  title: "문서 제목"              # 표지에 표시될 큰 제목
  project_name: "프로젝트명"      # 프로젝트 정보
  client: "고객사명"              # 고객사 정보
  date: "2025-11-17"            # 작성일자
  author: "작성자명"             # 작성자 이름
  version: "1.0"                # 문서 버전

sections:
  - title: "1단계 섹션"          # Heading 1 (새 페이지 시작)
    content: |                   # (선택) 일반 텍스트 내용
      여기에 설명 텍스트를 작성할 수 있습니다.
      여러 줄도 가능합니다.
    subsections:
      - title: "1-1 하위 섹션"   # Heading 2
        steps:                   # (선택) 단계별 설명 리스트
          - "첫 번째 단계입니다."
          - "두 번째 단계입니다."
      - title: "1-2 하위 섹션"   # Heading 2
        subsections:             # (선택) 더 깊은 계층 구조
          - title: "1-2-1"       # Heading 3
            steps:
              - "세부 단계입니다."
```

### 필드 설명

**metadata (표지 정보)**
- `title`: 문서 제목 - 표지에 28pt 크기로 표시
- `project_name`: 프로젝트명
- `client`: 고객사명
- `date`: 작성일자 (예: 2025-11-17)
- `author`: 작성자 이름
- `version`: 문서 버전 (예: 1.0)

**sections (본문 내용)**
- `title`: 섹션 제목 (필수)
- `content`: 일반 텍스트 내용 (선택, 여러 줄 가능)
- `steps`: 단계별 설명 리스트 (선택, List Paragraph 스타일)
- `subsections`: 하위 섹션 리스트 (선택, 재귀적 구조 지원)

## 프로젝트 파일 구조

```
wbs_app/
├── document_generator.py    # Word 문서 생성 핵심 모듈
├── create_document.py        # 메인 실행 스크립트
├── sample_data.yaml          # 샘플 데이터 (수정하여 사용)
├── requirements.txt          # Python 라이브러리 목록
├── README.md                 # 프로젝트 문서 (이 파일)
├── 산출물예시.docx           # 참고용 템플릿 문서
└── output.docx               # 생성된 Word 문서 (실행 후)
```

## 문서 스타일 규칙

### 페이지 설정
- **용지 크기**: A4 (8.27 × 11.69 inches)
- **여백**: 좌우 1.0", 상단 1.18", 하단 2.07"
- **푸터**: 중앙 정렬, "페이지번호 / 총페이지수"

### 제목 스타일 (맑은 고딕, 굵게)
- **Heading 1**: 16pt - 새 페이지 시작
- **Heading 2**: 13pt
- **Heading 3**: 12pt
- **Heading 4**: 11pt
- **Heading 5**: 10pt

### 본문 스타일
- **Normal**: 10pt, 맑은 고딕
- **List Paragraph**: 단계별 설명 (steps 필드)
- **목차**: 점선 리더, 하이퍼링크 (밑줄 없음)

## 사용 예시

### 기술 매뉴얼 생성 예시

`sample_data.yaml` 파일을 다음과 같이 수정:

```yaml
metadata:
  title: "시스템 관리자 매뉴얼"
  project_name: "통합 관리 시스템 v2.0"
  client: "샘플 기업"
  date: "2025-11-17"
  author: "시스템팀"
  version: "2.0"

sections:
  - title: "시스템 설치"
    subsections:
      - title: "사전 요구사항"
        steps:
          - "운영체제: Ubuntu 20.04 LTS 이상"
          - "메모리: 최소 8GB RAM"
          - "디스크 공간: 50GB 이상"
      - title: "설치 절차"
        steps:
          - "설치 파일을 다운로드합니다."
          - "압축을 해제합니다."
          - "install.sh 스크립트를 실행합니다."

  - title: "시스템 설정"
    content: |
      시스템 설치 후 초기 설정이 필요합니다.
      관리자 계정으로 로그인하여 다음 단계를 수행하세요.
    subsections:
      - title: "네트워크 설정"
        steps:
          - "관리 콘솔에 접속합니다."
          - "네트워크 메뉴를 선택합니다."
          - "IP 주소와 서브넷을 입력합니다."
```

실행:
```bash
python create_document.py
```

결과: `output.docx` 파일 생성

## 기술 세부사항

### 주요 기능 구현

1. **자동 필드 업데이트**
   - `settings.xml`에 `updateFields=true` 설정 추가
   - 문서를 열면 페이지 번호가 자동으로 업데이트됨

2. **하이퍼링크 목차**
   - OXML 요소로 북마크 생성
   - 하이퍼링크에서 밑줄 제거 (`u:none`, `color:000000`)
   - PAGEREF 필드로 페이지 번호 참조

3. **점선 리더**
   - 탭 스톱 설정 (`WD_TAB_ALIGNMENT.RIGHT`, `WD_TAB_LEADER.DOTS`)
   - 6인치 위치에 오른쪽 정렬 탭

4. **재귀적 섹션 구조**
   - `_add_section()` 메서드로 무제한 깊이 지원
   - 자동 번호 매기기 (1, 1.1, 1.1.1, ...)

## 참고사항

### 제한사항
- 이미지 삽입은 지원하지 않음 (Word에서 수동 추가 필요)
- 표(Table) 자동 생성은 미지원 (표지 정보 표 제외)

### YAML 작성 시 주의
- YAML 파일은 **UTF-8 인코딩**으로 저장
- 들여쓰기는 **스페이스 2칸** 사용 (탭 사용 금지)
- `content` 필드에서 여러 줄 작성 시 `|` 사용

### 페이지 구분 규칙
- **Heading 1**: 항상 새 페이지에서 시작
- **Heading 2-5**: 연속된 흐름으로 배치
- 페이지 구분을 줄이려면 Heading 1 사용을 최소화

## FAQ

**Q: 목차에 페이지 번호가 표시되지 않아요**
A: 문서를 열면 자동으로 업데이트됩니다. 수동 업데이트가 필요하면 목차 영역 우클릭 → "필드 업데이트"

**Q: 한글이 깨져 보여요**
A: YAML 파일을 UTF-8 인코딩으로 저장했는지 확인하세요.

**Q: 북마크 표시(검은 점)를 숨기고 싶어요**
A: Word 옵션 → 표시 → 책갈피 체크 해제

**Q: 출력 파일명을 변경하고 싶어요**
A: `create_document.py`의 `output_file` 변수를 수정하세요.

## 버전 히스토리

- **v1.0** (2025-11-17)
  - 초기 릴리스
  - 표지, 목차, 본문 자동 생성
  - 하이퍼링크 목차 (밑줄 없음)
  - 자동 필드 업데이트 기능
  - 재귀적 섹션 구조 지원

## 라이선스

MIT License
