# GitHub Pages 무료 고정 URL 배포 방법

이 앱은 Python 서버 없이 브라우저에서 엑셀을 직접 분석하는 정적 웹앱 구조입니다.
따라서 GitHub Pages에 무료로 배포할 수 있습니다.

## 업로드할 폴더

아래 폴더 안의 파일만 GitHub 저장소에 올리면 됩니다.

`github-pages-package`

포함 파일:
- index.html
- styles.css
- app.js
- exceljs.min.js
- microsoft-excel-logo.png
- .nojekyll

## GitHub 업로드 방법

1. GitHub에서 새 저장소를 만듭니다.
2. 저장소는 무료로 쓰려면 Public이어야 합니다.
3. `github-pages-package` 폴더 안의 파일들을 저장소 루트에 업로드합니다.
4. 저장소 첫 화면에 `index.html`, `app.js`, `styles.css`, `exceljs.min.js`가 바로 보여야 합니다.

## GitHub Pages 켜기

1. GitHub 저장소에서 Settings로 이동합니다.
2. 왼쪽 메뉴에서 Pages를 선택합니다.
3. Build and deployment에서 Source를 `Deploy from a branch`로 선택합니다.
4. Branch를 `main` / `/root`로 선택합니다.
5. Save를 누릅니다.
6. 잠시 후 Pages 주소가 생성됩니다.

예시 URL:
`https://사용자명.github.io/저장소명/`

## 보안 관련

- 엑셀 파일은 서버로 업로드되지 않고 사용자 브라우저 안에서 처리됩니다.
- ExcelJS도 외부 CDN이 아니라 `exceljs.min.js` 파일을 저장소에서 직접 불러옵니다.
- LOG 기록은 사용자 브라우저 localStorage에 저장됩니다.
- 같은 URL, 같은 브라우저에서 접속하면 PC를 껐다 켜도 일반적으로 유지됩니다.

## GitHub Pages에 필요 없는 파일

아래 파일은 GitHub Pages 배포에는 필요 없습니다.
- server.py
- requirements.txt
- render.yaml
- __pycache__