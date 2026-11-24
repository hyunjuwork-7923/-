# 테마의 한글(ea)·cs·Hang 스크립트 서체를 Pretendard로 통일한다
import sys, zipfile, re
zin = zipfile.ZipFile(sys.argv[1]); zout = zipfile.ZipFile(sys.argv[2], "w", zipfile.ZIP_DEFLATED)
for it in zin.infolist():
    d = zin.read(it.filename)
    if re.match(r"ppt/theme/theme\d+\.xml", it.filename):
        t = d.decode("utf8")
        t = re.sub(r'<a:ea typeface="[^"]*"/>', '<a:ea typeface="Pretendard"/>', t)
        t = re.sub(r'<a:cs typeface="[^"]*"/>', '<a:cs typeface="Pretendard"/>', t)
        t = re.sub(r'(<a:font script="Hang" typeface=")[^"]*(")', r'\1Pretendard\2', t)
        t = re.sub(r'<a:latin typeface="[^"]*"/>', '<a:latin typeface="Pretendard"/>', t)
        d = t.encode("utf8")
    zout.writestr(it, d)
zout.close()
