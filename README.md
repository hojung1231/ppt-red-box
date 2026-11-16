# PowerPoint 그리기 도구

Windows의 PowerPoint에서 동작하는 AutoHotkey v2 매크로입니다.

## 설치

1. [AutoHotkey v2](https://www.autohotkey.com/)를 설치합니다.
2. [ppt_red_box.ahk 다운로드](https://raw.githubusercontent.com/hojung1231/ppt-red-box/main/ppt_red_box.ahk)에서 파일을 받아 쓰기 가능한 폴더에 저장합니다.
3. 파일을 실행합니다. 단축키 목록은 시작할 때 표시되며, Windows 트레이 아이콘에서도 열 수 있습니다.

## 단축키

| 단축키 | 동작 |
| --- | --- |
| 휠 클릭 | PowerPoint 자유형 도구 시작 |
| 휠 드래그 | 사각형 미리보기 및 그리기 |
| Ctrl + Shift + A | 화살촉을 시작점, 끝점, 양쪽, 없음 순으로 변경 |
| Ctrl + Shift + C | 매크로로 만든 도형의 선 색을 빨강, 파랑, 노랑, 초록 순으로 변경 |
| Ctrl + Shift + X | 4 × 1.28cm 텍스트 상자 만들기. 나눔바른고딕 OTF 12pt |
| Ctrl + Shift + &#96; | 선 종류를 실선, 사각 점선, 파선 순으로 변경 |
| Ctrl + Shift + 휠 | 선 두께를 0.25pt씩 조절. 최소 0.25pt |

## 자동 업데이트

앱이 시작할 때와 실행 중 한 시간마다 이 저장소의 `main` 브랜치를 확인합니다. 앱의 **업데이트 확인** 버튼으로 즉시 확인할 수도 있습니다. 새 버전은 문법 검사 후 PowerPoint 창이 비활성화되면 적용되고 매크로만 다시 시작됩니다.

새 버전을 게시하려면 `ppt_red_box.ahk`의 `appVersion`을 올리고 해당 파일을 `main`에 커밋합니다.
