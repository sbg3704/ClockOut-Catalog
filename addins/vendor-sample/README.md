# JSON 보기 좋게(예제) (`vendor-sample`)

외부 라이브러리 묶음을 불러 쓰는 **예제 추가기능**이다. 카탈로그에는 올리지 않는다.

- 리본 칼퇴 탭: "JSON 보기 좋게"(2칸 들여쓰기), "JSON 한 줄로"
- 선언 `dependencies`에 `clockout-lib`(0.1.001)과 `clockout-vendor-data`(0.1.000)를 적었다. `build_addin.ps1`이 이것을 보고 연결부 `CoLibClient.bas`와 `CoVendorClient.bas`만 넣는다(lib·외부 원본 소스는 넣지 않는다).
- 시험: 묶음을 먼저 빌드한 뒤 이 예제를 빌드한다. 시험 때 `out\clockout-vendor-data.xlam`을 자동으로 열어 둔다.
```
powershell -ExecutionPolicy Bypass -File tools\build_addin.ps1 -Addin clockout-vendor-data -RunTests
powershell -ExecutionPolicy Bypass -File tools\build_addin.ps1 -Addin vendor-sample -RunTests
```
