# 🚀 Streamlit Cloud 배포 가이드

## 1. GitHub에 코드 업로드하기

### 1-1. Git 초기화 및 커밋

터미널에서 다음 명령어를 실행하세요:

```bash
# Git 초기화
git init

# 모든 파일 추가
git add .

# 커밋
git commit -m "우편봉투 인쇄 시스템 초기 버전"
```

### 1-2. GitHub 저장소 생성

1. https://github.com 접속 및 로그인
2. 오른쪽 상단의 "+" 버튼 클릭
3. "New repository" 선택
4. 저장소 이름 입력 (예: envelope-printing-system)
5. Public 또는 Private 선택
6. "Create repository" 클릭

### 1-3. GitHub에 푸시

GitHub에서 제공하는 명령어 사용:

```bash
# 원격 저장소 추가
git remote add origin https://github.com/본인계정명/저장소이름.git

# 메인 브랜치로 변경
git branch -M main

# 푸시
git push -u origin main
```

## 2. Streamlit Cloud에 배포하기

### 2-1. Streamlit Cloud 접속

1. https://share.streamlit.io 접속
2. GitHub 계정으로 로그인

### 2-2. 앱 배포

1. "New app" 버튼 클릭
2. 다음 정보 입력:
   - **Repository**: 방금 만든 GitHub 저장소 선택
   - **Branch**: main
   - **Main file path**: app.py
3. "Deploy!" 버튼 클릭

### 2-3. 배포 완료

- 몇 분 후 앱이 자동으로 배포됩니다
- 고유한 URL이 생성됩니다 (예: https://본인계정명-envelope-printing-system-app-xxxx.streamlit.app)
- 이 URL을 통해 어디서나 접속 가능합니다!

## 📋 필수 파일 체크리스트

배포 전 다음 파일들이 저장소에 포함되어 있는지 확인하세요:

- ✅ app.py - 메인 애플리케이션
- ✅ requirements.txt - Python 패키지 목록
- ✅ packages.txt - 시스템 패키지 (한글 폰트)
- ✅ number.xlsm - 정렬 기준 파일
- ✅ g.jpg - 로고 이미지
- ✅ .streamlit/config.toml - Streamlit 설정
- ✅ .gitignore - Git 제외 파일 목록

## 🔧 문제 해결

### 폰트 문제
- Streamlit Cloud는 Linux 환경이므로 Windows 폰트 대신 Nanum Gothic을 사용합니다
- packages.txt에 한글 폰트가 포함되어 있습니다

### 파일 크기 제한
- 업로드 파일 최대 크기: 200MB
- .streamlit/config.toml에서 설정됨

### 앱 업데이트
- GitHub에 코드를 푸시하면 자동으로 앱이 업데이트됩니다
```bash
git add .
git commit -m "업데이트 메시지"
git push
```

## 🎉 배포 완료!

이제 어디서나 웹 브라우저로 우편봉투 인쇄 시스템을 사용할 수 있습니다!

