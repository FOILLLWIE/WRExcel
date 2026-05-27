# 할인금액 계산기 URL 배포 방법

이 앱은 엑셀 파일을 서버에서 분석합니다.
그래서 GitHub Pages처럼 HTML만 올리는 방식으로는 정상 동작하지 않습니다.
Render, Railway, Fly.io처럼 Python 서버 실행이 가능한 서비스에 배포해야 합니다.

아래는 Render 기준 방법입니다.

## 1. GitHub 저장소 만들기

1. GitHub에 로그인합니다.
2. 새 저장소를 만듭니다.
3. 이 패키지의 system 폴더 안에 있는 파일들을 GitHub 저장소의 최상위 위치에 업로드합니다.

업로드해야 하는 파일:
- index.html
- styles.css
- app.js
- server.py
- requirements.txt
- render.yaml
- microsoft-excel-logo.png

중요:
release-package 폴더 전체를 올리는 것이 아니라, system 폴더 안의 파일들을 저장소 첫 화면에 바로 보이게 올리는 것이 가장 쉽습니다.

## 2. Render에서 웹 서비스 만들기

1. https://render.com 에 접속합니다.
2. 회원가입 또는 로그인합니다.
3. Dashboard에서 New + 버튼을 누릅니다.
4. Web Service를 선택합니다.
5. GitHub 저장소를 연결합니다.
6. 방금 만든 저장소를 선택합니다.
7. render.yaml 파일이 있으면 Render가 설정을 자동으로 읽습니다.
8. 배포를 시작합니다.
9. 배포가 끝나면 https://프로젝트명.onrender.com 형태의 URL이 생성됩니다.

## 3. Render 설정이 자동으로 안 잡힐 경우

아래처럼 직접 입력하면 됩니다.

Build Command:
pip install -r requirements.txt

Start Command:
HOST=0.0.0.0 python server.py

Environment:
Python

## 4. 배포 후 사용 방법

Render에서 생성된 URL을 다른 사람에게 공유하면 됩니다.
상대방은 별도 설치 없이 웹페이지에서 엑셀 파일을 업로드해서 사용할 수 있습니다.

## 5. LOG 저장에 대한 주의

LOG 기록은 서버가 아니라 사용자의 브라우저 localStorage에 저장됩니다.
따라서 같은 배포 URL, 같은 브라우저에서 접속하면 PC를 껐다 켜도 일반적으로 유지됩니다.

다만 아래 경우에는 기록이 이어지지 않을 수 있습니다.
- 브라우저 사이트 데이터 삭제
- 시크릿 모드 사용
- 다른 브라우저 사용
- 배포 URL 변경
- 엑셀 파일명 변경

## 6. 로컬 실행 파일 구조

release-package 폴더 바깥쪽에는 실행 파일만 있습니다.
시스템 파일은 system 폴더 안에 모아두었습니다.

로컬에서 실행할 때는 실행하기.vbs를 사용하면 CMD 창이 거의 보이지 않습니다.