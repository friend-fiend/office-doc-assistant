"""第二轮探测：pypdf 6.x 的页面压缩 / 旋转 / 叠加 / 加密权限枚举。"""
import io


def check(name, fn):
    try:
        r = fn()
        print(f"[OK]   {name} -> {r}")
    except Exception as e:
        print(f"[FAIL] {name} -> {type(e).__name__}: {e}")


from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas


def make_pdf(pages=3) -> bytes:
    b = io.BytesIO()
    c = canvas.Canvas(b)
    for i in range(pages):
        c.drawString(72, 720, f"Page {i + 1} hello world content " * 3)
        c.showPage()
    c.save()
    return b.getvalue()


SRC = make_pdf()


def t1():
    """加入 writer 之后再压缩内容流。"""
    r = PdfReader(io.BytesIO(SRC))
    w = PdfWriter()
    for p in r.pages:
        w.add_page(p)
    for p in w.pages:
        p.compress_content_streams(level=9)
    b = io.BytesIO()
    w.write(b)
    return f"{len(b.getvalue())} bytes"


check("compress_content_streams 在 writer 内", t1)


def t2():
    """compress_identical_objects 参数。"""
    r = PdfReader(io.BytesIO(SRC))
    w = PdfWriter()
    w.append(r)
    w.compress_identical_objects(remove_identicals=True, remove_orphans=True)
    b = io.BytesIO()
    w.write(b)
    return f"{len(b.getvalue())} bytes"


check("compress_identical_objects", t2)


def t3():
    """旋转 + transfer_rotation_to_content。"""
    r = PdfReader(io.BytesIO(SRC))
    w = PdfWriter()
    for p in r.pages:
        p.rotate(90)
        w.add_page(p)
    b = io.BytesIO()
    w.write(b)
    out = PdfReader(io.BytesIO(b.getvalue()))
    return {"rotation": out.pages[0].rotation,
            "mediabox": (float(out.pages[0].mediabox.width), float(out.pages[0].mediabox.height))}


check("page.rotate(90)", t3)


def t4():
    """叠加覆盖层 merge_page。"""
    r = PdfReader(io.BytesIO(SRC))
    w = PdfWriter()
    for p in r.pages:
        p.transfer_rotation_to_content()
        W, H = float(p.mediabox.width), float(p.mediabox.height)
        ob = io.BytesIO()
        oc = canvas.Canvas(ob, pagesize=(W, H))
        oc.setFillAlpha(0.3)
        oc.setFont("Helvetica", 40)
        oc.saveState()
        oc.translate(W / 2, H / 2)
        oc.rotate(45)
        oc.drawString(-50, 0, "WATERMARK")
        oc.restoreState()
        oc.showPage()
        oc.save()
        ov = PdfReader(io.BytesIO(ob.getvalue()))
        p.merge_page(ov.pages[0], over=True)
        w.add_page(p)
    b = io.BytesIO()
    w.write(b)
    PdfReader(io.BytesIO(b.getvalue()))
    return f"{len(b.getvalue())} bytes"


check("merge_page 叠加覆盖层", t4)


def t5():
    """encrypt：permissions_flag 传 int 是否可行（RC4-128）。"""
    r = PdfReader(io.BytesIO(SRC))
    w = PdfWriter()
    w.append(r)
    w.encrypt(user_password="pw123", owner_password="pw123",
              permissions_flag=0b0000000000000100 | 0b0000000000000001,
              algorithm="RC4-128")
    b = io.BytesIO()
    w.write(b)
    chk = PdfReader(io.BytesIO(b.getvalue()))
    enc = chk.is_encrypted
    ok = chk.decrypt("pw123")
    return {"encrypted": enc, "decrypt_ok": bool(ok), "bytes": len(b.getvalue())}


check("encrypt RC4-128 (int 权限位)", t5)


def t6():
    """encrypt：algorithm=None + use_128bit（老式调用）。"""
    return "skip"
    return "skip"


def t7():
    """UserAccessPermissions 枚举可用性。"""
    from pypdf.constants import UserAccessPermissions
    return {"PRINT": int(UserAccessPermissions.PRINT),
            "EXTRACT": int(UserAccessPermissions.EXTRACT),
            "MODIFY": int(UserAccessPermissions.MODIFY),
            "ADD_OR_MODIFY": int(UserAccessPermissions.ADD_OR_MODIFY)}


check("UserAccessPermissions 枚举", t7)


def t8():
    """枚举方式设置权限。"""
    from pypdf.constants import UserAccessPermissions
    perm = UserAccessPermissions.PRINT | UserAccessPermissions.PRINT_TO_REPRESENTATION
    r = PdfReader(io.BytesIO(SRC))
    w = PdfWriter()
    w.append(r)
    w.encrypt("pw", "pw", permissions_flag=perm, algorithm="RC4-128")
    b = io.BytesIO()
    w.write(b)
    chk = PdfReader(io.BytesIO(b.getvalue()))
    return {"encrypted": chk.is_encrypted, "decrypt": bool(chk.decrypt("pw"))}


check("encrypt 使用枚举权限", t8)


def t9():
    """AES-128 在无 cryptography 时的报错表现。"""
    r = PdfReader(io.BytesIO(SRC))
    w = PdfWriter()
    w.append(r)
    try:
        w.encrypt("pw", "pw", algorithm="AES-128")
        b = io.BytesIO()
        w.write(b)
        return f"AES 可用 {len(b.getvalue())} bytes"
    except Exception as e:
        return f"AES 不可用 -> {type(e).__name__}: {e}"


check("AES-128 无 cryptography", t9)


def t10():
    """extract_text() 可用性（用于文档对比）。"""
    r = PdfReader(io.BytesIO(SRC))
    return repr((r.pages[0].extract_text() or "")[:60])


check("extract_text", t10)
