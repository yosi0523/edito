# My Calendar 앱 - Firebase 설정 가이드

## 1단계: Firebase 프로젝트 만들기

1. https://console.firebase.google.com 접속
2. 구글 계정으로 로그인
3. "프로젝트 추가" 클릭
4. 프로젝트 이름: `my-calendar` 입력
5. Google 애널리틱스는 "사용 안 함" 선택
6. "프로젝트 만들기" 클릭

## 2단계: 웹앱 등록

1. 프로젝트 홈에서 </> (웹) 아이콘 클릭
2. 앱 닉네임: `My Calendar` 입력
3. "Firebase 호스팅도 설정" 체크
4. "앱 등록" 클릭
5. 나오는 설정값(firebaseConfig)을 메모해두세요!

## 3단계: 인증(Authentication) 설정

1. 왼쪽 메뉴 "Authentication" 클릭
2. "시작하기" 클릭
3. "Google" 로그인 제공업체 클릭
4. "사용 설정" 활성화
5. 프로젝트 지원 이메일 선택
6. "저장" 클릭

## 4단계: Firestore 데이터베이스 설정

1. 왼쪽 메뉴 "Firestore Database" 클릭
2. "데이터베이스 만들기" 클릭
3. 위치: `asia-northeast3` (서울) 선택
4. "프로덕션 모드에서 시작" 선택
5. "만들기" 클릭
6. 생성 후 "규칙" 탭에서 아래 내용으로 교체:

```
rules_version = '2';
service cloud.firestore {
  match /databases/{database}/documents {
    match /users/{userId}/{document=**} {
      allow read, write: if request.auth != null && request.auth.uid == userId;
    }
  }
}
```

7. "게시" 클릭

## 5단계: Storage 설정

1. 왼쪽 메뉴 "Storage" 클릭
2. "시작하기" 클릭
3. "프로덕션 모드에서 시작" 선택
4. 위치: `asia-northeast3` 선택
5. 생성 후 "Rules" 탭에서 아래 내용으로 교체:

```
rules_version = '2';
service firebase.storage {
  match /b/{bucket}/o {
    match /users/{userId}/{allPaths=**} {
      allow read, write: if request.auth != null && request.auth.uid == userId;
    }
  }
}
```

6. "게시" 클릭

## 6단계: 설정값 넣기

1. `app.js` 파일을 열어주세요
2. 맨 위의 `firebaseConfig` 부분을 2단계에서 메모한 값으로 교체하세요

## 7단계: 호스팅 (인터넷에 올리기)

### Firebase Hosting 사용 (무료)

1. 터미널에서:
```
npm install -g firebase-tools
firebase login
cd C:/Users/NHN/my-calendar-app
firebase init hosting
firebase deploy
```

2. 배포 완료되면 `https://프로젝트이름.web.app` 주소가 나옵니다

### 모바일에서 앱처럼 설치

- **아이폰**: Safari에서 접속 → 공유 버튼 → "홈 화면에 추가"
- **안드로이드**: Chrome에서 접속 → 메뉴(⋮) → "홈 화면에 추가"
