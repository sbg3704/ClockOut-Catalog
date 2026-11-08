#!/usr/bin/env python3
# 기능: addins/_common/의 공개 함수(Co*.bas, ClockOutSettings.bas)를 읽어 공개 진입점 클래스 CoLibApi.cls를 자동으로 만든다.
#       추가기능은 Application.Run("'clockout-lib.xlam'!CoLib_Api")로 이 개체를 얻어 늦은 바인딩으로 부른다.
# 사용법: python3 tools/gen_colibapi.py          -> addins/_common/CoLibApi.cls 다시 쓰기
#         python3 tools/gen_colibapi.py --check  -> 지금 CoLibApi.cls가 최신이 아니면 종료 코드 1(함수 더한 뒤 잊었는지 확인)
#   lib에 공개 함수를 더하면 이 스크립트를 다시 돌려 같은 PR에 CoLibApi.cls를 함께 올린다.
#   사용자 정의 형을 쓰는 함수(CoGuard_Begin/End, CoFitMath_CalcBox)와 ParamArray 함수(CoSetting_Choice)는 SPECIAL에 손으로 쓴 감싸기를 쓴다.
# 의존: 파이썬 3 표준 라이브러리만.
import re, sys, pathlib

ROOT = pathlib.Path(__file__).resolve().parent.parent / "addins" / "_common"
OUT = ROOT / "CoLibApi.cls"
OBJ_TYPES = {"Shape", "ShapeRange", "Range", "Collection", "Object", "Worksheet", "Workbook"}
SKIP_MODULES = {"CoLibEntry"}

HEADER = '''VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CoLibApi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'==============================================================
' 기능: 칼퇴 공용 라이브러리(clockout-lib)의 공개 진입점 개체. 모든 Co 모듈의 공개 함수를 같은 이름의 메서드로 감싸 넘긴다.
'       Co 모듈은 Option Private Module이라 다른 통합 문서에서 직접 부를 수 없으므로 이 개체를 거친다.
' 사용법(추가기능 쪽, 늦은 바인딩):
'   Dim lib As Object
'   Set lib = Application.Run("'clockout-lib.xlam'!CoLib_Api")
'   If lib.CoRange_GetSelection(ws, rng, why) Then ...      ' ByRef 인자도 그대로 통한다
'   tok = lib.CoGuard_Begin() : ... : lib.CoGuard_End tok   ' 상태는 개체 안에 보관(토큰)
'   v = lib.CoFitMath_CalcBox(...)                          ' Variant 배열(left, top, width, height), 못 맞추면 Empty
'   s = lib.FilterPictures                                  ' 상수는 읽기 속성으로
' 주의: 이 파일은 tools/gen_colibapi.py가 만든다. 손으로 고치지 말고 lib 함수를 고친 뒤 스크립트를 다시 돌린다.
' 의존: addins/_common의 모든 Co 모듈과 ClockOutSettings.bas
'==============================================================
Option Explicit

Private mStates() As CoAppState
Private mStateCount As Long

'--- 상수 읽기 ------------------------------------------------
Public Property Get MsgTitle() As String: MsgTitle = CO_MSG_TITLE: End Property
Public Property Get FilterAll() As String: FilterAll = CO_FILTER_ALL: End Property
Public Property Get FilterPictures() As String: FilterPictures = CO_FILTER_PICTURES: End Property
Public Property Get FilterExcel() As String: FilterExcel = CO_FILTER_EXCEL: End Property
Public Property Get FilterCsv() As String: FilterCsv = CO_FILTER_CSV: End Property
Public Property Get MmToPtFactor() As Double: MmToPtFactor = CO_MM_TO_PT: End Property
'''

SPECIAL = {
"CoGuard_Begin": '''' 상태(CoAppState)는 개체 경계를 넘지 못하므로 개체 안에 보관하고 토큰(1부터)을 돌려준다. CoGuard_End에 그대로 넘긴다.
Public Function CoGuard_Begin(Optional ByVal noScreen As Boolean = True, _
                              Optional ByVal noEvents As Boolean = False, _
                              Optional ByVal manualCalc As Boolean = False, _
                              Optional ByVal noAlerts As Boolean = False) As Long
    mStateCount = mStateCount + 1
    ReDim Preserve mStates(1 To mStateCount)
    mStates(mStateCount) = CoGuard.CoGuard_Begin(noScreen, noEvents, manualCalc, noAlerts)
    CoGuard_Begin = mStateCount
End Function
''',
"CoGuard_End": '''' CoGuard_Begin이 준 토큰의 상태로 되돌린다. 모르는 토큰이거나 이미 되돌렸으면 아무것도 안 한다.
Public Sub CoGuard_End(ByVal token As Long)
    If token < 1 Or token > mStateCount Then Exit Sub
    CoGuard.CoGuard_End mStates(token)
End Sub
''',
"CoFitMath_CalcBox": '''' 결과를 Variant 배열(0:left, 1:top, 2:width, 3:height)로 돌려준다. 맞출 수 없으면 Empty.
Public Function CoFitMath_CalcBox(ByVal areaLeft As Double, ByVal areaTop As Double, _
                                  ByVal areaW As Double, ByVal areaH As Double, _
                                  ByVal marginPt As Double, _
                                  ByVal shapeW As Double, ByVal shapeH As Double, _
                                  ByVal keepRatio As Boolean, ByVal alignCenter As Boolean) As Variant
    Dim b As CoFitBox
    b = CoFitMath.CoFitMath_CalcBox(areaLeft, areaTop, areaW, areaH, marginPt, shapeW, shapeH, keepRatio, alignCenter)
    If b.Ok Then CoFitMath_CalcBox = Array(b.Left, b.Top, b.Width, b.Height)
End Function
''',
"CoSetting_Choice": '''' 선택지 설정. ParamArray는 넘겨줄 수 없어 읽은 값을 여기서 options와 견준다(동작은 CoSetting_Choice와 같다).
Public Function CoSetting_Choice(ByVal addinId As String, ByVal key As String, ByVal defaultValue As String, _
                                 ParamArray options() As Variant) As String
    Dim s As String, i As Long
    s = CoSetting.CoSetting_Choice(addinId, key, defaultValue)   ' 선택지 없이 부르면 읽은 값 그대로
    CoSetting_Choice = defaultValue
    If UBound(options) < LBound(options) Then CoSetting_Choice = s: Exit Function
    For i = LBound(options) To UBound(options)
        If s = CStr(options(i)) Then CoSetting_Choice = s: Exit Function
    Next i
End Function
''',
}

SIG = re.compile(r'^Public (Function|Sub) (\w+)\((.*)\)(?: As (\w+))?\s*$')

def logical_lines(text):
    buf = ""
    for raw in text.replace("\r\n", "\n").split("\n"):
        line = raw.rstrip()
        if line.endswith(" _"):
            buf += line[:-2].strip() + " "
            continue
        yield (buf + line.strip()) if buf else raw.strip()
        buf = ""

def split_params(p):
    out, depth, cur, inq = [], 0, "", False
    for ch in p:
        if ch == '"': inq = not inq
        if ch == "(" and not inq: depth += 1
        if ch == ")" and not inq: depth -= 1
        if ch == "," and depth == 0 and not inq:
            out.append(cur.strip()); cur = ""
        else:
            cur += ch
    if cur.strip(): out.append(cur.strip())
    return out

def arg_name(param):
    words = [w for w in param.split("=")[0].split() if w not in ("Optional", "ByVal", "ByRef", "ParamArray")]
    return words[0].rstrip("()")

def build():
    parts, seen = [HEADER], set()
    for f in sorted(ROOT.glob("*.bas")):
        mod = f.stem
        if mod in SKIP_MODULES: continue
        text = f.read_text(encoding="utf-8")
        m = re.search(r'Attribute VB_Name = "(\w+)"', text)
        if m: mod = m.group(1)
        funcs = []
        for line in logical_lines(text):
            s = SIG.match(line)
            if not s: continue
            kind, name, params, ret = s.groups()
            if name.startswith("Test_"): continue
            funcs.append((kind, name, params, ret))
        if not funcs: continue
        parts.append(f"\n'--- {mod} " + "-" * max(1, 56 - len(mod)) + "\n")
        for kind, name, params, ret in funcs:
            if name in seen: sys.exit(f"이름이 겹칩니다: {name}")
            seen.add(name)
            if name in SPECIAL:
                parts.append(SPECIAL[name]); continue
            if "ParamArray" in params or (ret and ret not in OBJ_TYPES | {"String", "Long", "Double", "Boolean", "Variant", "Date", "Integer"}):
                sys.exit(f"{mod}.{name}: 자동으로 감쌀 수 없는 형입니다. SPECIAL에 손으로 더하세요.")
            args = ", ".join(arg_name(p) for p in split_params(params))
            call = f"{mod}.{name}({args})" if kind == "Function" else (f"{mod}.{name} {args}" if args else f"{mod}.{name}")
            sig = f"Public {kind} {name}({params})" + (f" As {ret}" if ret else "")
            if kind == "Sub":
                body = f"    {call}\n"
            elif ret in OBJ_TYPES:
                body = f"    Set {name} = {call}\n"
            else:
                body = f"    {name} = {call}\n"
            parts.append(f"{sig}\n{body}End {kind}\n")
    missing = set(SPECIAL) - seen
    if missing: sys.exit(f"SPECIAL에 있는데 lib에 없는 함수: {sorted(missing)}")
    return "".join(parts).replace("\r\n", "\n").replace("\n", "\r\n")

if __name__ == "__main__":
    new = build()
    if "--check" in sys.argv:
        old = OUT.read_bytes().decode("utf-8") if OUT.exists() else ""
        if old != new:
            print("CoLibApi.cls가 최신이 아닙니다: python3 tools/gen_colibapi.py 를 돌리세요."); sys.exit(1)
        print("CoLibApi.cls 최신"); sys.exit(0)
    OUT.write_bytes(new.encode("utf-8"))
    print("만듦:", OUT.relative_to(ROOT.parent.parent), new.count("\nPublic "), "개 공개 멤버")
