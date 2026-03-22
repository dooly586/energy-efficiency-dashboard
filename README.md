# 에너지 효율화 사업 대시보드 (Energy Efficiency Dashboard)

이 프로젝트는 고효율 보일러의 에너지 절감 효과를 시각화하고 투자 수익성을 분석하는 웹 애플리케이션입니다. React와 Vite를 기반으로 구축되었으며, 사용자가 데이터를 업로드하고 분석 결과를 리포트로 확인할 수 있는 기능을 제공합니다.

## 🚀 시작하기

프로젝트를 로컬 환경에서 실행하려면 아래 단계를 따르세요.

### 📋 사전 준비 사항

- [Node.js](https://nodejs.org/) (최신 LTS 버전 권장)
- npm (Node.js 설치 시 함께 설치됨)

### 🛠️ 설치 및 실행

1.  **의존성 패키지 설치:**
    ```bash
    npm install
    ```

2.  **개발 서버 실행:**
    ```bash
    npm run dev
    ```
    실행 후 터미널에 표시되는 로컬 주소(예: `http://localhost:5173`)를 브라우저에서 엽니다.

3.  **프로젝트 빌드 (배포용):**
    ```bash
    npm run build
    ```

4.  **빌드 결과 미리보기:**
    ```bash
    npm run preview
    ```

## 🛠️ 주요 기술 스택

- **Frontend:** React 19, Vite 6
- **Styling:** Vanilla CSS (모던 UI 디자인)
- **Charts:** Recharts
- **Icons:** Lucide React
- **Utilities:** xlsx (엑셀 데이터 처리), html2canvas (이미지 캡처)

## 📁 프로젝트 구조

- `src/App.jsx`: 메인 로직 및 레이아웃
- `src/index.css`: 전역 스타일 및 디자인 시스템
- `public/`: 정적 자원
