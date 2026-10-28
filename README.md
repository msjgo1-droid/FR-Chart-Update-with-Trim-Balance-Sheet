# FR Chart TRIM 자동 채우기

Balance Sheet(스타일별) 여러 개 + FR(Friday Review) 차트 1개를 업로드하면,
각 스타일의 TRIM 영역을 Balance Sheet 데이터로 자동 채운 FR 파일을 만들어주는 Streamlit 웹앱입니다.

## 파일 구성
- `app.py` — Streamlit UI
- `engine.py` — 실제 엑셀 처리 로직
- `requirements.txt` — 필요한 파이썬 패키지

## 로컬에서 실행해보기
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Streamlit Community Cloud에 배포하기

1. **GitHub 저장소 만들기**
   - github.com 에서 New repository 클릭
   - 이름 예: `fr-trim-autofill` (원하는 이름으로)
   - Public 또는 Private 선택 (Streamlit Cloud 무료 플랜은 Public 저장소만 지원)
   - 이 3개 파일(`app.py`, `engine.py`, `requirements.txt`)을 저장소에 업로드
     - GitHub 웹사이트에서 "Add file → Upload files" 로 그냥 드래그해서 올려도 됩니다

2. **Streamlit Community Cloud 연결**
   - https://share.streamlit.io 접속 후 GitHub 계정으로 로그인
   - "New app" 클릭
   - 방금 만든 저장소 선택, 브랜치는 `main`, Main file path는 `app.py` 입력
   - Deploy 클릭

3. 몇 분 후 `https://[앱이름].streamlit.app` 형태의 링크가 생성됩니다. 이전에 쓰시던
   `pi-hoka-shipping-documents` 사이트와 같은 방식입니다.

## 사용법
1. Balance Sheet 파일들(.xlsx, 스타일당 1개)을 업로드
2. FR 차트 파일(.xlsx) 1개 업로드
3. "처리 시작" 클릭
4. 결과 확인 후 업데이트된 FR 파일 다운로드
