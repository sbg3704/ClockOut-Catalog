# 추가기능 빌드 스크립트 (tools/build_addin.ps1)

`addins/<id>/`의 소스를 Excel에서 컴파일해 `.xlam`과 설치용 `.zip`까지 한 번에 만든다.
(인코딩 변환, 가져오기, 컴파일 확인, 저장, 묶기를 모두 자동으로 한다. 한글판 VBA 편집기가 읽는 CP949 변환도 포함.)

## 준비물
- Windows용 **Excel 데스크톱판** (웹판·Mac판·LibreOffice 안 됨)
- Excel 설정 한 번: 파일 > 옵션 > 보안 센터 > 보안 센터 설정 > 매크로 설정 >
  **'VBA 프로젝트 개체 모델에 대한 액세스 신뢰'** 체크. Windows 계정마다 따로 저장되니 계정이 다르면 각각 켠다.
  (스크립트는 이 설정을 바꾸지 않는다. 꺼져 있으면 안내를 띄우고 멈춘다.)

## 실행 (저장소 루트에서)
```
powershell -ExecutionPolicy Bypass -File tools\build_addin.ps1 -Addin image-fit
```
- `-Addin` : `addins\` 아래 폴더 이름(= 추가기능 id)
- `-RunTests` : `Test_`로 시작하는 시험 함수도 실행(창은 뜨지 않음, 하나라도 실패하면 xlam/zip을 만들지 않음)
- `-OutDir <폴더>` : 출력 폴더 바꾸기(기본 `out\`)
- 공용 라이브러리(`addins\_common`)만 컴파일·시험하려면 `-Addin _common`. 결과는 `out\clockout-lib.xlam`/`.zip`이고 선언·ThisWorkbook 코드는 필요 없다.

Excel 창은 뜨지 않고(보이지 않게 실행), 키보드 입력이나 화면 조작은 하지 않는다. 스크립트가 띄운 Excel만 끝에 종료하며, 이미 열어 둔 Excel은 건드리지 않는다.

## 결과물
저장소 루트의 `out\` (git에는 올라가지 않음)
- `out\<id>.xlam` : 추가기능 본체
- `out\<id>.zip` : 칼퇴 [추가기능 설치]에서 고르는 패키지 (맨 위에 `.xlam`, `<id>.clockout.json`, `README.md`가 있으면 그것, `src\` 아래에 .bas 전부(UTF-8 원본)와 `ThisWorkbook_*.txt`)

함께 넣는 소스는 선언(`<id>.clockout.json`)으로 정한다. `addins\<id>\ThisWorkbook_*.txt`가 있으면 ThisWorkbook 코드로 넣는다.
| 선언 | 함께 넣는 것 |
|---|---|
| `"type": "library"` (외부 라이브러리 묶음 `clockout-vendor-*`) | `vendor.txt`에 적은 `vendor\vba\` 원본만. zip에 `licenses\<라이브러리>\LICENSE`와 `licenses\SOURCES.md`도 넣음 |
| `dependencies`에 `clockout-lib`(bundled 아님) | `addins\_client\CoLibClient.bas` |
| `dependencies`에 `clockout-vendor-*` | `addins\_client\CoVendorClient.bas` |
| `clockout-lib`이 `bundled: true`이거나 선언·dependencies 없음(예전 방식) | `addins\_common\*.bas` 전부 |

- `-RunTests` 때 선언 dependencies의 라이브러리가 `out\<id>.xlam`으로 있으면 읽기 전용으로 열어 두고 시험한다. 그래서 라이브러리를 먼저 빌드한다(예: `-Addin clockout-vendor-data` 다음 `-Addin vendor-sample`).

## 리본 규칙 (2026-11-09)
- 단추·메뉴처럼 콜백(onAction)을 받는 컨트롤은 **일반 `id`**로 만든다. `idQ`(이름공간 붙은 id)로 만들면 엑셀이 클릭을 그 이름공간의 COM 추가기능으로 보내 VBA 콜백이 불리지 않는다(리본은 보이는데 단추가 무반응, 0.1.003까지의 문제).
- 여러 추가기능이 함께 쓰는 탭(`ck:main`)과 그룹만 `idQ`로 둔다.
- 콜백은 추가기능마다 `CoRibbon_Click_<id>`로 이름이 바뀌어 들어간다(빌드가 처리).

## 자주 나는 오류
| 메시지 | 원인과 해결 |
|---|---|
| `Excel이 VBA 프로젝트 접근을 막고 있습니다` | 위 '준비물'의 신뢰 설정을 켠다 |
| `Excel(데스크톱판)을 실행할 수 없습니다` | Excel 미설치 또는 설치가 깨짐. Excel이 있는 PC에서 실행 |
| `컴파일 오류가 있습니다 … 모듈 'X' N번째 줄` | 그 모듈의 그 줄을 고친다. 변수 이름이 VBA 예약어·내장 함수(`scale`, `inStr` 등)와 같으면 흔히 난다 |
| `CP949로 쓸 수 없는 글자가 있습니다` | 소스에 이모지 등 한글판 VBA가 못 읽는 글자가 있다. 지운다 |
| `같은 이름의 모듈 파일이 addins\<id> 와 addins\_common 에 모두 있습니다` | 한쪽 이름을 바꾸거나 지운다 |
| `이 시스템에서 스크립트를 실행할 수 없습니다`(실행 정책) | 위 실행 줄처럼 `-ExecutionPolicy Bypass`를 붙인다 |

종료 코드: 0 성공, 1 입력·일반 오류, 2 Excel 없음, 3 VBA 접근 신뢰 꺼짐, 4 컴파일 오류, 5 시험 실패.
