# AI 자동 채점 (AI Auto-Grading)

Google Apps Script 기반의 AI 자동 채점 도구입니다. Google Classroom 제출물을 OpenAI, Claude 등의 AI API를 사용하여 자동으로 채점합니다.

## 기능

- **Google Classroom 연동**: Classroom 과제 제출물 자동 수집
- **다중 AI API 지원**: OpenAI Chat Completions, OpenAI Responses, Claude Messages API 지원
- **스프레드시트 기반 관리**: Google 스프레드시트에서 채점 큐, 제출물, 결과를 관리
- **자동 큐 처리**: 설정한 간격으로 채점 큐를 자동 처리
- **재시도 메커니즘**: API 호출 실패 시 자동 재시도
- **커스텀 채점 기준**: API 규격에 따른 유연한 채점 프롬프트 설정

## 설치 방법

1. 이 리포지토리의 파일을 Google Apps Script 프로젝트에 복사합니다.
2. `Code.gs`와 `sidebar.html`을 각각 스크립트 편집기에 추가합니다.
3. 필요한 Google 스프레드시트의 시트를 생성합니다:
   - `AI채점_작업` (작업 큐)
   - `AI채점_제출물` (제출물 목록)
   - `AI채점_결과` (채점 결과)

## 사용 방법

1. 스프레드시트에서 **AI 자동 채점** 메뉴를 클릭합니다.
2. **채점 사이드바 열기**를 선택하여 사이드바를 표시합니다.
3. API 설정에서 사용할 AI 서비스와 모델을 선택합니다.
4. Google Classroom 과제를 선택하거나 직접 제출물을 업로드합니다.
5. 채점을 시작하면 자동으로 결과가 스프레드시트에 기록됩니다.

## 설정

| 설정 항목 | 설명 | 기본값 |
|----------|------|--------|
| API 규격 | 사용할 AI API 형식 | Chat Completions |
| 모델 | 사용할 AI 모델 | gpt-4o-mini |
| 최대 출력 토큰 | AI 응답 최대 길이 | 65536 |
| 재시도 횟수 | API 호출 실패 시 재시도 횟수 | 5회 |

## 라이선스

MIT License