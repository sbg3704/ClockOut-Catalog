#!/usr/bin/env python3
# 기능: addins/ 아래 .bas(UTF-8)를 VBA 편집기가 읽는 CP949+CRLF로 바꿔 지정 폴더에 저장한다.
# 사용법: python3 tools/vba_encode.py <출력폴더>   (예: python3 tools/vba_encode.py out_vba)
#         출력 폴더의 .bas를 VBA 편집기 > 파일 가져오기로 넣는다. 저장소에는 UTF-8 원본만 둔다.
# 의존: 파이썬 3 표준 라이브러리만.
import sys, pathlib

def main():
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    out = pathlib.Path(sys.argv[1]); out.mkdir(parents=True, exist_ok=True)
    root = pathlib.Path(__file__).resolve().parent.parent / "addins"
    for src in sorted(root.rglob("*.bas")):
        text = src.read_text(encoding="utf-8").replace("\r\n", "\n").replace("\n", "\r\n")
        try:
            data = text.encode("cp949")
        except UnicodeEncodeError as e:
            sys.exit(f"{src}: CP949로 쓸 수 없는 글자가 있습니다: {e.object[e.start:e.end]!r}")
        (out / src.name).write_bytes(data)
        print("변환:", src.relative_to(root.parent), "->", out / src.name)

main()
