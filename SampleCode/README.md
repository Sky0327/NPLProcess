# Samil NPL 평가 시스템

부실채권(NPL) 평가를 위한 웹 기반 데이터 수집 및 리포트 생성 시스템

## 기술 스택

- React 18
- Material-UI (MUI)
- React Router
- Vite

## 설치 및 실행

```bash
# 의존성 설치
npm install

# 개발 서버 실행
npm run dev

# 빌드
npm run build

# 빌드 미리보기
npm run preview
```

## 프로젝트 구조

```
src/
├── components/          # 재사용 가능한 컴포넌트
│   ├── Layout/         # 레이아웃 컴포넌트
│   └── SmartNPL1/      # Smart_NPL1 관련 컴포넌트
├── pages/              # 페이지 컴포넌트
│   ├── Home/          # 홈 페이지
│   ├── SmartNPL1/     # 데이터 수집 및 처리
│   └── SmartNPL2/     # 리포트 생성 및 내보내기
├── services/           # API 서비스
├── data/              # Mock 데이터
├── theme.js           # Material-UI 테마 설정
├── App.jsx            # 메인 앱 컴포넌트
└── main.jsx           # 진입점
```

## 주요 기능

### Smart_NPL1 - 데이터 수집 및 처리

1. **초기화 및 설정**
   - 보고서명 설정
   - 프로젝트 ID 설정
   - API 인증 정보 설정

2. **병렬 데이터 조회**
   - 등기조회
   - 공시지가조회
   - 법원경매조회
   - 인포케어 통계
   - 인포케어 통합
   - 실거래가조회 (국토/밸류맵)

3. **중간 처리**
   - KB시세조회
   - 인포케어 사례
   - 거리계산

4. **리포트 생성**
   - 담보물정보 리포트
   - 감정평가 리포트
   - 경매정보 리포트
   - 낙찰통계/사례 리포트
   - 실거래사례 리포트

### Smart_NPL2 - 리포트 생성 및 내보내기

- 최종 리포트 생성
- PDF 파일명 변환
- Excel 내보내기
- 파일 초기화

## 브랜드 컬러

- Primary: #aa3142 (짙은 버건디)
- Secondary: #6c757d
- Background: #f8f9fa (밝은 테마)

## 개발 참고사항

- 현재 모든 API 호출은 Mock 데이터를 반환합니다.
- 실제 API 연동 시 `src/services/apiService.js`를 수정하세요.
- Mock 데이터는 `src/data/mockData.js`에서 관리합니다.

