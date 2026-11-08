# 칼퇴 공용 VBA 라이브러리 (`addins/_common`)

> 이 폴더는 앱 저장소(sbg3704/ClockOut)의 `addins/_common`에서 이 저장소로 옮겨 왔다(2026-11-08). 문서 속 `docs/02`·`docs/03`·PR 번호는 앱 저장소 기준이다.

추가기능을 만들 때마다 같은 도우미를 다시 짜지 않도록 모아 둔 공용 모듈이다. 새 매크로는 **먼저 여기서 찾아 쓰고, 없을 때만 새로 만든다.** 필요한 모듈의 `.bas`만 골라 추가기능에 가져오면 된다.

**의존 방향은 한쪽뿐이다: 추가기능 → 공용 lib.** 이 폴더의 모듈은 어떤 추가기능 폴더(`addins/<이름>/`)의 파일에도 기대지 않고, 이 폴더만으로 컴파일된다.

## 관리 원칙 (2026-11-08, 사용자 지시)
- 이 폴더는 **계정2가 상시 관리**한다. 사용자나 계정1이 필요한 기능을 요청하면 계정2가 함수를 더하거나 고친다.
- 추가기능 매크로는 이 lib에 의존한다. 같은 일을 하는 함수를 매크로마다 새로 선언하지 않는다. 새 매크로에 lib에 없는 도우미가 필요하면 매크로 안에 만들지 말고 `협업/요청/`으로 lib에 더해 달라고 알린다.
- lib는 추가기능에 의존하지 않는다(한쪽 방향).
- 요청은 한 일마다 별도 PR로 올리고, 완료는 PR#4(계정1 우편함)에 리뷰 댓글로 알린다. 무게(+1/+10/+100)만 적는다.

## 모듈
| 모듈 | 하는 일 | 의존 | 순수 계산 |
|---|---|---|---|
| `ClockOutSettings.bas` | 표준 읽기 함수 `ClockOut_GetSetting(id, key, 기본값)` (`docs/03`의 표준 이름이라 Co 접두어 없음) | 없음 | |
| `CoSetting.bas` | 설정 한 줄 읽기: 숫자(min/max 제한)·mm→pt·참거짓·글자·선택지 확인 | ClockOutSettings, CoNum | |
| `CoGuard.bas` | ScreenUpdating·이벤트·계산·상태 표시줄 저장/복원 가드, 오류 안내 글, 안내·확인 창 | 없음 | |
| `CoRange.bas` | 활성 워크시트·선택 범위 안전 검사(행·열 전체, 숨김, 병합, 보호, 잠긴 칸), 마지막 행·열, 범위 → 2차원 배열 | 없음 | |
| `CoShape.bas` | 선택 도형 얻기, 사진 안전하게 넣기·지우기, 칸에 맞추기, 범위 안 도형 | CoFitMath | |
| `CoFitMath.bas` | 칸 안에 맞출 위치·크기 계산 `CoFitMath_CalcBox`(여백·비율·정렬) | 없음 | ○ `Test_CoFitMath` |
| `CoFile.bas` | 파일·여러 파일·폴더·저장 위치 고르기(취소 = ""), 있는지 확인, 폴더 만들기, UTF-8 읽기·쓰기 | CoText | |
| `CoText.bas` | 빈칸 판정, 보이지 않는 글자 정리(`CoText_CleanInvisible`), 시작·끝 비교, 나누기, 안전한 파일 이름, 경로 잇기·나누기, 열 이름↔번호 | 없음 | ○ `Test_CoText` |
| `CoNum.bas` | 범위 제한, mm·cm·pt 변환, 지역 설정과 무관한 숫자 읽기, 보통 반올림, 단위 절사(`CoNum_TruncTo`), 한글 금액(`CoNum_ToKoreanWon`), 0 나누기 방지 | 없음 | ○ `Test_CoNum` |

## 이름 규칙
- 모듈은 `Co<분야>`, 공개 함수는 `Co<분야>_<동사·명사>`(예: `CoRange_GetSelection`). 여러 추가기능을 같이 켜도 이름이 겹치지 않고, 이름만 봐도 어느 모듈인지 안다.
- 공개 상수는 `CO_` 대문자(예: `CO_FILTER_PICTURES`, `CO_MM_TO_PT`).
- 모든 Co 모듈은 `Option Private Module`이다. 시트 함수 목록과 Alt+F8 매크로 목록에 안 보이고, 같은 추가기능 안에서만 쓴다. 시험 매크로는 직접 실행 창에서 이름을 쳐서 돌린다.
- 지역 변수·인자 이름으로 `name`·`out`·`scale`·`inStr`·`base`처럼 VBA 예약어나 내장 함수와 겹칠 수 있는 낱말을 쓰지 않는다(한글판 Excel에서 컴파일 오류가 난 적 있음, PR#7). `fileName`·`outText`·`fitScale`처럼 붙여 쓴다.
- 검사 함수는 안내 창을 직접 띄우지 않고 `reason`(이유 글)을 돌려준다. 띄울지는 매크로가 정한다.
- 실패해도 오류를 밖으로 던지지 않고 False·""·기본값을 돌려준다(오프라인·안전 원칙, `docs/02` 3장 4번).

## 매크로 뼈대 (복사해서 시작)
```vb
Private Const ADDIN_ID As String = "my-addin"   ' .clockout.json의 id와 같게

Public Sub MyAddin_Run()
    Dim st As CoAppState, msg As String, why As String
    Dim ws As Worksheet, rng As Range, margin As Double

    If Not CoRange_GetSelection(ws, rng, why) Then CoGuard_Notice why: Exit Sub
    If Not CoRange_CanEdit(rng) Then CoGuard_Notice "보호된 셀이 있습니다.": Exit Sub
    margin = CoSetting_Number(ADDIN_ID, "margin", 1#, 0, 20)   ' JSON의 default·min·max와 같게

    st = CoGuard_Begin()
    On Error GoTo EH
    ' ... 작업 ...
    CoGuard_End st
    Exit Sub
EH:
    msg = CoGuard_ErrText("작업 중")   ' 맨 먼저: 다음 줄들이 Err를 지운다
    CoGuard_End st
    CoGuard_Notice msg
End Sub
```

## 가져오는 법
1. `python3 tools/vba_encode.py out_vba` — UTF-8 원본을 VBA 편집기가 읽는 CP949+CRLF로 바꾼다. `addins/` 아래 `.bas`를 모두 바꾸므로 이 폴더도 같이 나온다.
2. VBA 편집기 > 파일 가져오기로 필요한 모듈과 그 의존 모듈만 넣는다(위 표의 "의존" 칸).
3. `CoText`·`CoNum`·`CoFitMath`를 넣었으면 직접 실행 창에서 `Test_CoText`, `Test_CoNum`, `Test_CoFitMath`를 돌려 "모두 통과"를 확인한다.

## 함수를 더할 때
1. 이미 있는지 이 표와 각 모듈 맨 위 주석을 먼저 본다. 비슷한 것이 있으면 인자를 늘려 고친다.
2. Office 개체를 안 쓰는 계산은 `CoText`·`CoNum`(또는 새 순수 모듈)에 두고 `Test_` 시험 줄을 함께 더한다(`docs/02` 3장 3번).
3. 모듈 맨 위 한국어 주석의 사용법 줄과 이 표를 같이 고친다.
4. 이미 쓰는 공개 함수의 이름·인자 순서는 바꾸지 않는다(다른 추가기능이 깨진다). 바꿔야 하면 새 이름으로 더한다.

## 확인 범위
- 클라우드(Linux)에서 작성했다. CP949로 바꿀 수 있는지만 확인했다. **VBA 컴파일·Excel 실행·시험 매크로는 아직 못 돌렸다.** Windows Excel에서 이 폴더의 모듈을 모두 가져와 컴파일(디버그 > 컴파일)하고 시험 매크로 3개를 돌려야 한다.
- `ClockOutSettings.bas`와 `tools/vba_encode.py`는 PR#7에 있는 것과 한 글자도 다르지 않은 복사본이다(합칠 때 충돌 없음). 앞으로는 이 폴더 쪽만 고친다.
- `CoFitMath`는 image-fit의 `ImageFitMath.CalcFitBox`와 같은 규칙이다. image-fit은 이 PR이 합쳐진 뒤 공용 모듈을 쓰도록 바꾸고 자기 복사본을 지운다(후속).
