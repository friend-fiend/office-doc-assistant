"""API 探测脚本：验证各依赖库当前版本的真实可用接口。"""
import inspect
import io


def check(name, fn):
    try:
        r = fn()
        print(f"[OK]   {name} -> {r}")
    except Exception as e:
        print(f"[FAIL] {name} -> {type(e).__name__}: {e}")


import docx
import openpyxl
import pdf2image
import PIL
import pptx
import py7zr
import pypdf
import reportlab

print("versions:", "pypdf", pypdf.__version__, "| reportlab", reportlab.Version,
      "| Pillow", PIL.__version__, "| py7zr", py7zr.__version__)


def t1():
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
    return pdfmetrics.stringWidth("第 1 页 / 共 10 页", "STSong-Light", 10)


check("reportlab UnicodeCIDFont STSong-Light", t1)


def t2():
    from reportlab.pdfgen import canvas
    b = io.BytesIO()
    c = canvas.Canvas(b)
    c.setFillAlpha(0.3)
    c.setStrokeAlpha(0.3)
    c.saveState()
    c.translate(10, 10)
    c.rotate(45)
    c.drawString(0, 0, "x")
    c.restoreState()
    c.showPage()
    c.save()
    return f"{len(b.getvalue())} bytes"


check("reportlab setFillAlpha + rotate", t2)


def t3():
    from pypdf import PdfWriter
    w = PdfWriter()
    return {"compress_identical_objects": hasattr(w, "compress_identical_objects"),
            "append": hasattr(w, "append"),
            "add_outline_item": hasattr(w, "add_outline_item"),
            "encrypt_sig": str(inspect.signature(w.encrypt))}


check("pypdf PdfWriter API", t3)


def t4():
    from pypdf import PdfReader
    from reportlab.pdfgen import canvas
    b = io.BytesIO()
    c = canvas.Canvas(b)
    c.drawString(72, 72, "hello")
    c.showPage()
    c.save()
    r = PdfReader(io.BytesIO(b.getvalue()))
    p = r.pages[0]
    p.compress_content_streams()
    return {"transfer_rotation_to_content": hasattr(p, "transfer_rotation_to_content"),
            "rotate": hasattr(p, "rotate"), "merge_page": hasattr(p, "merge_page"),
            "mediabox": (float(p.mediabox.width), float(p.mediabox.height))}


check("pypdf Page API", t4)


def t5():
    try:
        import cryptography
        return "cryptography present " + cryptography.__version__
    except Exception as e:
        return f"no cryptography ({e})"


check("pypdf AES support", t5)


def t6():
    from PIL import Image
    return {"Image.LANCZOS": hasattr(Image, "LANCZOS"), "Resampling": hasattr(Image, "Resampling")}


check("Pillow resampling", t6)


def t7():
    return {"FILTER_LZMA2": hasattr(py7zr, "FILTER_LZMA2"),
            "SevenZipFile sig": str(inspect.signature(py7zr.SevenZipFile.__init__)),
            "exceptions.PasswordRequired": hasattr(py7zr.exceptions, "PasswordRequired")}


check("py7zr API", t7)


def t8():
    return str(inspect.signature(pdf2image.convert_from_path))[:700]


check("pdf2image convert_from_path", t8)

check("openpyxl Workbook", lambda: openpyxl.Workbook().sheetnames)
check("python-docx Document", lambda: type(docx.Document()).__name__)
check("python-pptx Presentation", lambda: type(pptx.Presentation()).__name__)


def t10():
    from pdf2image.exceptions import PDFInfoNotInstalledError, PDFPageCountError  # noqa: F401
    return "ok"


check("pdf2image exceptions", t10)


def t11():
    from PIL import Image
    im = Image.new("RGB", (100, 50), (255, 0, 0))
    b = io.BytesIO()
    im.save(b, "PDF", save_all=True, append_images=[], resolution=200.0, quality=90, optimize=True)
    return f"{len(b.getvalue())} bytes"


check("Pillow 多图存 PDF", t11)
