# CountLocales

Excel 파일의 다국어 텍스트 글자 수 및 단어 수 분석 도구

## 📋 개요

CountLocales는 Excel 파일(.xlsx, .xlsm, .csv)에서 다국어 텍스트의 글자 수와 단어 수를 분석하는 도구입니다. 여러 언어를 지원하며, 각 열별로 상세한 통계를 제공합니다.

## ✨ 주요 기능

### 글자 수 분석 (Character Count)
- **지원 언어**: 한국어, 영어, 중국어, 일본어, 태국어, 러시아어, 숫자, 특수문자
- **분석 항목**:
  - 실제 글자 수 (중복 포함)
  - 고유 텍스트 기준 글자 수 (시트별)
  - 고유 텍스트 기준 글자 수 (폴더 전체)
  - 셀 주소 정보
  - 셀 개수 통계

### 단어 수 분석 (Word Count)
- **지원 언어**: 한국어, 영어, 스페인어, 프랑스어, 독일어, 포르투갈어, 이탈리아어, 러시아어, 중국어, 일본어, 베트남어, 태국어, 인도네시아어, 터키어
- **자연어 처리**:
  - 한국어: Kiwi 형태소 분석기
  - 영어/유럽 언어: spaCy
  - 중국어: jieba
  - 일본어: Stanza
- **분석 항목**:
  - 실제 단어 수 (중복 포함)
  - 고유 텍스트 기준 단어 수 (시트별)
  - 고유 텍스트 기준 단어 수 (폴더 전체)
  - 셀 주소 정보
  - 셀 개수 통계

### 기타 기능
- 하위 폴더 포함 자동 검색
- 다국어 UI 지원 (한국어/영어)
- 빈 열 자동 감지 및 중단 (연속 20개)
- 대용량 데이터 처리 (임시 파일 활용)

## 🚀 설치 방법

### 요구사항
- Python 3.7 이상

### 기본 설치
```bash
pip install -r requirements.txt
```

### 선택적 설치 (자연어 처리 향상)
```bash
# spaCy 언어 모델 (단어 수 분석용)
python -m spacy download en_core_web_sm
python -m spacy download es_core_news_sm
python -m spacy download fr_core_news_sm
python -m spacy download de_core_news_sm
python -m spacy download pt_core_news_sm
python -m spacy download it_core_news_sm
python -m spacy download ru_core_news_sm
```

## 📖 사용 방법

### 실행 파일 사용 (권장)
1. `dist/main.exe` 실행
2. 언어 선택 (한국어/English)
3. 분석 방식 선택 (글자 수 분석/단어 수 분석)
4. 프로그램이 실행된 폴더의 모든 Excel 파일을 자동으로 분석

### Python 스크립트 실행
```bash
python main.py
```

### 실행 파일 빌드
```bash
pyinstaller --onefile --console main.py
```

## 📊 출력 형식

프로그램은 다음과 같은 Excel 보고서를 생성합니다:

### 글자 수 분석 보고서
- **Summary_real**: 실제 글자 수 (중복 포함)
- **Summary_unique_for_Sheet**: 시트별 고유 텍스트 기준 글자 수
- **Summary_unique_for_Folder**: 폴더 전체 고유 텍스트 기준 글자 수
- **Summary_cell_address**: 각 글자가 포함된 셀 주소
- **Summary_cells**: 각 언어별 셀 개수

### 단어 수 분석 보고서
- **Words_real**: 실제 단어 수 (중복 포함)
- **Words_unique_for_Sheet**: 시트별 고유 텍스트 기준 단어 수
- **Words_unique_for_Folder**: 폴더 전체 고유 텍스트 기준 단어 수
- **Words_cell_address**: 각 단어가 포함된 셀 주소
- **Words_cells**: 각 언어별 셀 개수

## 📁 프로젝트 구조

```
countlocales/
├── main.py              # 메인 진입점
├── count_chars.py       # 글자 수 분석 모듈
├── count_words.py       # 단어 수 분석 모듈
├── translations.py      # 다국어 번역 딕셔너리
├── requirements.txt    # Python 패키지 의존성
└── README.md           # 프로젝트 문서
```

## 🔧 기술 스택

- **Python 3.7+**
- **pandas**: Excel 파일 처리
- **openpyxl**: Excel 파일 읽기/쓰기
- **langdetect**: 언어 자동 감지
- **kiwipiepy**: 한국어 형태소 분석
- **spaCy**: 다국어 자연어 처리
- **jieba**: 중국어 단어 분리
- **stanza**: 일본어 자연어 처리
- **tqdm**: 진행 상황 표시

## 📝 라이선스

MIT License

## 🤝 기여

이슈 및 풀 리퀘스트를 환영합니다!

## 📧 문의

프로젝트 관련 문의사항이 있으시면 이슈를 등록해주세요.

