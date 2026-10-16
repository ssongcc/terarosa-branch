# 지점 원두 우체국 접수 앱

지점 원두 발주서(.xlsx)를 올리면 지점별·박스별 우체국 택배 접수파일을 만들어 주는 Streamlit 앱입니다.

## 파일 구성

| 파일 | 역할 |
|---|---|
| `app.py` | 화면 (업로드 → 결과 확인 → 다운로드) |
| `core.py` | 박스 계산·접수파일 생성 로직 |
| `requirements.txt` | 필요한 패키지 |
| `.streamlit/config.toml` | 테라로사 색상 테마 |

## 배포 (Streamlit Community Cloud, 공개 저장소 방식)

주소록(지점 연락처)은 GitHub에 올리지 않고 Streamlit **Secrets**에만 넣습니다.

1. GitHub에 **Public** 저장소 생성 → 아래 5개만 업로드 (**주소 xlsx는 올리지 않음**)
   `app.py`, `core.py`, `requirements.txt`, `README.md`, `.streamlit/config.toml`
2. share.streamlit.io → Create app → 저장소 / `main` / `app.py`
3. Advanced settings → Secrets 에 `Secrets_붙여넣기용.toml` 내용 전체 붙여넣기 (비밀번호 수정) → Deploy
4. 앱 링크 + 비밀번호를 팀원에게 공유

## 운영

- **지점 주소 변경/신규 지점**: 앱 Settings → Secrets 에서 해당 `[[addresses]]` 블록 수정·추가 → Save
  - 급할 때는 앱 사이드바에서 주소 xlsx를 올리면 그 작업에만 적용
- **발주서 지점 이름과 주소록 이름이 다를 때**: `core.py`의 `ALIAS`에 추가
  - 예: `"현대송도": "현대아울렛송도"` (발주서 헤더 → 주소록 수취인명에 포함된 글자)
  - 매칭이 안 되면 앱에 빨간 경고로 표시되고 접수파일에서 빠짐
- **제외 지점 변경**: `core.py`의 `EXCLUDE`
- **박스 기준 변경**: `core.py`의 `BOX_KG = 15`

## 계산 규칙

1. 어센틱에스프레소블렌드 1kg: 지점별 kg → 15kg 기준 박스(올림), 박스마다 고르게 분배 (20kg → 10+10, 40kg → 14+13+13)
2. 그 외 품명 뒤 중량(g/kg) 표기 원두: 지점별 합산 → 15kg 기준 박스(올림)
3. 수취인명 뒤 `(박스번호/총박스)` 표기, 어센틱 박스는 품목에 `어센틱에스프레소블렌드1kg 10kg` 표기
4. 강릉 지점(본점·경포·임당·사천·신라모노) 제외
