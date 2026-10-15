# -*- coding: utf-8 -*-
"""PasteCompress —— 图片 / PDF 轻量压缩器

用法：
    双击 启动压缩器.bat，或  python PasteCompress.py

核心能力：
    * Ctrl+V 直接粘贴剪贴板里的截图 / 图片文件 / PDF 文件（可一次多个）
    * 五档预设 + 手动微调（最长边、质量、格式、灰度、目标大小），实时预估压缩后体积
    * 图片即时预览，PDF 逐页预览，支持「原图 / 压缩后 / 并排」对比
    * 一键导出，或直接复制回剪贴板（图片=位图，PDF/文件=文件）
    * PDF 优先用 pikepdf 做结构保留压缩（文字矢量不丢），可选 pypdfium2 整页栅格化极限压缩
    全部本地处理，不联网、不上传。
"""

from __future__ import annotations

import base64
import collections
import io
import json
import os
import queue
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import traceback
from dataclasses import dataclass, field, replace
from pathlib import Path

APP_NAME = "压一压"
APP_SUBTITLE = "图片 / PDF 轻量压缩器"
APP_VERSION = "1.2.0"
FROZEN = bool(getattr(sys, "frozen", False))
if FROZEN:                                   # 打包成 exe 后：资源在 _MEIPASS，输出在 exe 旁边
    APP_DIR = Path(sys.executable).resolve().parent
    BUNDLE_DIR = Path(getattr(sys, "_MEIPASS", APP_DIR))
else:
    APP_DIR = Path(__file__).resolve().parent
    BUNDLE_DIR = APP_DIR

# ---------------------------------------------------------------- 依赖引导
if not FROZEN:
    _LIBS = APP_DIR / "libs"
    if _LIBS.is_dir() and str(_LIBS) not in sys.path:
        sys.path.insert(0, str(_LIBS))

if sys.stdout is None:                       # 打包成无控制台 exe 时这两个是 None
    sys.stdout = open(os.devnull, "w")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w")

from PIL import Image, ImageFile, ImageOps  # noqa: E402

ImageFile.LOAD_TRUNCATED_IMAGES = True
try:
    from PIL import ImageTk
    HAS_IMAGETK = True
except Exception:                                     # pragma: no cover
    HAS_IMAGETK = False

if int(getattr(Image, "__version__", "0").split(".")[0]) < 11:
    try:                                              # Pillow<11 需要插件才有 AVIF
        import pillow_avif  # noqa: F401
    except Exception:
        pass


def _soft_import(name):
    try:
        return __import__(name)
    except Exception:
        return None


pikepdf = _soft_import("pikepdf")
pdfium = _soft_import("pypdfium2")
HAS_PIKEPDF = pikepdf is not None
HAS_PDFIUM = pdfium is not None

IMAGE_EXTS = {".jpg", ".jpeg", ".jpe", ".jfif", ".png", ".webp", ".bmp", ".gif",
              ".tif", ".tiff", ".avif", ".heic", ".heif", ".ico", ".jp2",
              ".j2k", ".ppm", ".pgm", ".tga", ".dds"}
PDF_EXTS = {".pdf"}

_MB = 1024.0 * 1024.0


# ---------------------------------------------------------------- 小工具
def human_size(n) -> str:
    """把字节数变成好读的字符串。"""
    try:
        n = float(n)
    except Exception:
        return "-"
    if n < 1024:
        return f"{n:.0f} B"
    if n < _MB:
        return f"{n / 1024:.1f} KB"
    if n < _MB * 1024:
        return f"{n / _MB:.2f} MB"
    return f"{n / (_MB * 1024):.2f} GB"


def kind_of(path_or_name: str) -> str:
    """"image" / "pdf" / "" """
    ext = Path(str(path_or_name)).suffix.lower()
    if ext in PDF_EXTS:
        return "pdf"
    if ext in IMAGE_EXTS:
        return "image"
    return ""


def unique_path(path: Path) -> Path:
    """避免覆盖已有文件：name_1.ext, name_2.ext ..."""
    if not path.exists():
        return path
    stem, suffix, parent = path.stem, path.suffix, path.parent
    for i in range(1, 1000):
        cand = parent / f"{stem}_{i}{suffix}"
        if not cand.exists():
            return cand
    return parent / f"{stem}_{int(time.time())}{suffix}"


def temp_root() -> Path:
    d = Path(tempfile.gettempdir()) / "PasteCompress"
    d.mkdir(parents=True, exist_ok=True)
    return d


def new_temp_file(suffix: str = "", prefix: str = "pc_") -> Path:
    fd, name = tempfile.mkstemp(suffix=suffix, prefix=prefix, dir=str(temp_root()))
    os.close(fd)
    return Path(name)


def cleanup_old_temp(days: float = 1.0) -> None:
    """清理上次运行留下的临时文件。"""
    try:
        cutoff = time.time() - days * 86400
        for p in temp_root().iterdir():
            try:
                if p.is_file() and p.stat().st_mtime < cutoff:
                    p.unlink()
            except Exception:
                pass
    except Exception:
        pass


def image_has_alpha(im: Image.Image) -> bool:
    if im.mode in ("RGBA", "LA", "PA", "La"):
        return True
    if im.mode == "P":
        return "transparency" in im.info
    return False


# ---------------------------------------------------------------- 剪贴板（Windows）
CREATE_NO_WINDOW = 0x08000000


def _run_ps(script: str, timeout: int = 30) -> str:
    """用 Windows PowerShell(STA) 执行脚本，返回输出文本。"""
    exe = shutil.which("powershell.exe") or shutil.which("powershell") or shutil.which("pwsh.exe")
    if not exe:
        raise RuntimeError("找不到 PowerShell")
    full = "[Console]::OutputEncoding=[Text.Encoding]::UTF8\n" + script
    enc = base64.b64encode(full.encode("utf-16-le")).decode("ascii")
    proc = subprocess.run([exe, "-NoProfile", "-NonInteractive", "-STA", "-EncodedCommand", enc],
                          capture_output=True, timeout=timeout, creationflags=CREATE_NO_WINDOW)
    out = proc.stdout.decode("utf-8", "replace")
    err = proc.stderr.decode("utf-8", "replace")
    if proc.returncode != 0 and not out.strip():
        raise RuntimeError(err.strip() or f"PowerShell 退出码 {proc.returncode}")
    return (out + err).strip()


def _ps_quote(s) -> str:
    return "'" + str(s).replace("'", "''") + "'"


def clipboard_get():
    """读取剪贴板。

    返回 (kind, payload)：
        ("image", PIL.Image)  —— 截图 / 复制的位图
        ("files", [Path,...]) —— 在资源管理器里复制的文件
        ("none",  None)
    """
    # 1) 走 PIL（CF_DIB 位图 + CF_HDROP 文件列表）
    try:
        from PIL import ImageGrab
        got = ImageGrab.grabclipboard()
        if isinstance(got, Image.Image):
            return "image", got
        if isinstance(got, list) and got:
            return "files", [Path(p) for p in got if isinstance(p, str) and p]
    except Exception:
        pass
    # 2) 兜底：PowerShell 读文件列表 / 文本路径
    try:
        out = _run_ps(
            "Add-Type -AssemblyName System.Windows.Forms\n"
            "$f=[System.Windows.Forms.Clipboard]::GetFileDropList()\n"
            "if($f -and $f.Count -gt 0){ foreach($i in $f){ Write-Output (\"F:\" + $i) } exit }\n"
            "$t=[System.Windows.Forms.Clipboard]::GetText()\n"
            "if($t){ Write-Output (\"T:\" + $t) }\n")
    except Exception:
        return "none", None
    files = []
    for line in out.splitlines():
        line = line.strip()
        if line.startswith("F:"):
            files.append(Path(line[2:].strip().strip('"')))
        elif line.startswith("T:"):
            cand = Path(line[2:].strip().strip('"').strip())
            try:
                if cand.is_file():
                    files.append(cand)
            except Exception:
                pass
    files = [p for p in files if p.is_file()] if files else []
    if files:
        return "files", files
    return "none", None


def _dib_blob(im: "Image.Image") -> bytes:
    """把 PIL 图像打包成 CF_DIB（BMP 去文件头；BGR 顺序、自底向上、4 字节对齐）。

    统一用 24 位并在白底上合成：这样粘到任何程序里颜色都一致
    （32 位带 alpha 的 DIB 在不同程序里表现不一致）。
    需要保留透明通道请用「导出」保存文件。
    """
    import struct
    src_im = _flatten(im)                       # 透明 -> 白底
    r, g, b = src_im.split()
    pixels_bgr = Image.merge("RGB", (b, g, r)).tobytes()
    w, h = src_im.size
    src_stride = w * 3
    dst_stride = ((w * 24 + 31) // 32) * 4
    pad = b"\x00" * (dst_stride - src_stride)
    rows = []
    for y in range(h - 1, -1, -1):              # DIB 自底向上
        rows.append(pixels_bgr[y * src_stride:(y + 1) * src_stride] + pad)
    body = b"".join(rows)
    header = struct.pack("<IiiHHIIiiII", 40, w, h, 1, 24, 0, len(body), 2835, 2835, 0, 0)
    return header + body


def _clip_set_dib(im: "Image.Image") -> None:
    """用 Windows API 直接把位图写进剪贴板（不依赖 PowerShell）。"""
    import ctypes
    from ctypes import wintypes
    u = ctypes.windll.user32
    k = ctypes.windll.kernel32
    CF_DIB = 8
    GMEM_MOVEABLE = 0x0002
    u.OpenClipboard.argtypes = [wintypes.HWND]
    u.CloseClipboard.argtypes = []
    u.EmptyClipboard.argtypes = []
    u.SetClipboardData.restype = ctypes.c_void_p
    u.SetClipboardData.argtypes = [wintypes.UINT, ctypes.c_void_p]
    k.GlobalAlloc.restype = ctypes.c_void_p
    k.GlobalAlloc.argtypes = [wintypes.UINT, ctypes.c_size_t]
    k.GlobalLock.restype = ctypes.c_void_p
    k.GlobalLock.argtypes = [ctypes.c_void_p]
    k.GlobalUnlock.argtypes = [ctypes.c_void_p]

    blob = _dib_blob(im)
    opened = False
    for _try in range(8):                     # 剪贴板常被别的程序短暂占用，重试几次
        if u.OpenClipboard(None):
            opened = True
            break
        time.sleep(0.12)
    if not opened:
        raise RuntimeError("剪贴板被其它程序占用（打不开），请关掉剪贴板增强/截图工具后重试")
    try:
        u.EmptyClipboard()
        h = k.GlobalAlloc(GMEM_MOVEABLE, len(blob))
        if not h:
            raise RuntimeError("GlobalAlloc 失败")
        ptr = k.GlobalLock(h)
        if not ptr:
            raise RuntimeError("GlobalLock 失败")
        ctypes.memmove(ptr, blob, len(blob))
        k.GlobalUnlock(h)
        if not u.SetClipboardData(CF_DIB, h):
            raise RuntimeError("SetClipboardData 失败")
    finally:
        u.CloseClipboard()


def _clip_seq() -> int:
    """剪贴板序号：任何程序写入剪贴板，这个号都会变（不需要打开剪贴板）。"""
    try:
        import ctypes
        from ctypes import wintypes
        u = ctypes.windll.user32
        u.GetClipboardSequenceNumber.restype = wintypes.DWORD
        return int(u.GetClipboardSequenceNumber())
    except Exception:
        return 0


def clipboard_has_image(timeout: float = 1.0) -> bool:
    """读回校验：剪贴板里现在有没有位图（剪贴板忙时多试几次）。"""
    import time as _t
    deadline = _t.time() + max(0.0, timeout)
    while True:
        try:
            from PIL import ImageGrab
            got = ImageGrab.grabclipboard()
            if isinstance(got, Image.Image) and got.size[0] > 0:
                return True
        except Exception:
            pass
        if _t.time() >= deadline:
            return False
        _t.sleep(0.15)


def clipboard_has_files(timeout: float = 1.0) -> bool:
    """读回校验：剪贴板里现在有没有文件列表（剪贴板忙时多试几次）。"""
    import time as _t
    deadline = _t.time() + max(0.0, timeout)
    while True:
        try:
            from PIL import ImageGrab
            got = ImageGrab.grabclipboard()
            if isinstance(got, list) and got:
                if any(Path(x).is_file() for x in got if isinstance(x, str)):
                    return True
        except Exception:
            pass
        if _t.time() >= deadline:
            return False
        _t.sleep(0.15)


def clipboard_set_image(path: Path):
    """把图片放进剪贴板：先 ctypes(CF_DIB)，不行再退 PowerShell。

    返回 (方式, 是否读回验证成功)。
    """
    errors = []
    wrote = None
    for how, action in (
            ("CF_DIB", lambda: _clip_set_dib(load_image(Path(path).read_bytes()))),
            ("System.Drawing", lambda: _clip_set_image_ps(path))):
        before = _clip_seq()
        try:
            action()
            wrote = wrote or how
        except Exception as exc:
            errors.append(f"{how}: {exc}")
            continue
        changed = _clip_seq() != before
        if changed or clipboard_has_image():
            return how, True                       # 序号变了 = 确实写进去了
        errors.append(f"{how}: 写入后剪贴板没有变化（可能被剪贴板工具占用）")
    if wrote:
        return wrote, False
    raise RuntimeError("；".join(errors))


def _clip_set_image_ps(path: Path) -> None:
    script = (
        "Add-Type -AssemblyName System.Windows.Forms\n"
        "Add-Type -AssemblyName System.Drawing\n"
        f"$img=[System.Drawing.Image]::FromFile({_ps_quote(path)})\n"
        "$bmp=New-Object System.Drawing.Bitmap($img)\n"
        "[System.Windows.Forms.Clipboard]::SetImage($bmp)\n"
        "$bmp.Dispose(); $img.Dispose()\n"
        "Write-Output 'OK'\n")
    out = _run_ps(script)
    if "OK" not in out:
        raise RuntimeError(f"写入剪贴板失败: {out[:200]}")


def clipboard_set_files(paths, verify: bool = False):
    """把文件放进剪贴板（资源管理器/微信/邮件里可直接 Ctrl+V 粘贴文件）。"""
    lines = ["Add-Type -AssemblyName System.Windows.Forms",
             "$c=New-Object System.Collections.Specialized.StringCollection"]
    for p in paths:
        lines.append(f"[void]$c.Add({_ps_quote(Path(p).resolve())})")
    lines += [
        "$d=New-Object System.Windows.Forms.DataObject",
        "$d.SetFileDropList($c)",
        "$ms=New-Object System.IO.MemoryStream",
        "$ms.Write([byte[]](5,0,0,0),0,4)",
        "[void]$d.SetData('Preferred DropEffect',$ms)",
        "[System.Windows.Forms.Clipboard]::SetDataObject($d,$true)",
        "Write-Output 'OK'"]
    before = _clip_seq()
    out = _run_ps("\n".join(lines))
    if "OK" not in out:
        raise RuntimeError(f"写入剪贴板失败: {out[:200]}")
    if not verify:
        return None
    verified = (_clip_seq() != before) or clipboard_has_files()
    return verified

# ---------------------------------------------------------------- 结果与参数
@dataclass
class Result:
    ok: bool = False
    kind: str = ""                 # "image" / "pdf"
    data: bytes = b""
    ext: str = ""
    note: str = ""
    error: str = ""
    stats: dict = field(default_factory=dict)
    warnings: list = field(default_factory=list)

    @property
    def size(self) -> int:
        return len(self.data)


@dataclass
class Settings:
    # 通用
    preset: str = "balanced"
    target_mode: bool = False
    target_kb: int = 500
    # 图片
    img_max_dim: int = 0           # 0 = 不缩放
    img_quality: int = 78
    img_format: str = "auto"       # auto / jpeg / webp / png / avif / keep
    img_gray: bool = False
    keep_meta: bool = False
    keep_if_bigger: bool = True
    # PDF
    pdf_img_max: int = 1600        # 内嵌图片最长边上限，0 = 不缩放
    pdf_quality: int = 74          # 0 = 不重编码内嵌图片
    pdf_gray: bool = False
    pdf_raster: bool = False       # 整页栅格化（文字变图片）
    pdf_raster_dpi: int = 110
    pdf_objstm: bool = True        # 用对象流重新打包

    def cloned(self) -> "Settings":
        return replace(self)


PRESETS = [
    ("lossless", "无损", dict(
        img_max_dim=0, img_quality=100, img_format="keep",
        pdf_img_max=0, pdf_quality=0, pdf_raster=False)),
    ("high", "高质量", dict(
        img_max_dim=2560, img_quality=86, img_format="auto",
        pdf_img_max=2400, pdf_quality=84, pdf_raster=False)),
    ("balanced", "平衡", dict(
        img_max_dim=1920, img_quality=78, img_format="auto",
        pdf_img_max=1600, pdf_quality=74, pdf_raster=False)),
    ("extreme", "极限", dict(
        img_max_dim=1280, img_quality=62, img_format="auto",
        pdf_img_max=1100, pdf_quality=60, pdf_raster=True, pdf_raster_dpi=96)),
]
PRESET_NAMES = {k: n for k, n, _ in PRESETS}


def preset_settings(key: str, base: "Settings | None" = None) -> Settings:
    """生成某档预设的参数（保留用户勾选的通用项）。"""
    s = base.cloned() if base else Settings()
    for k, _name, over in PRESETS:
        if k == key:
            for kk, vv in over.items():
                setattr(s, kk, vv)
            break
    s.preset = key
    return s


_LANCZOS = getattr(getattr(Image, "Resampling", Image), "LANCZOS", 1)


# ---------------------------------------------------------------- 图片引擎
def load_image(data: bytes) -> Image.Image:
    """打开图片：处理多帧、透明调色板、EXIF 方向，并保住原始 format 信息。"""
    im = Image.open(io.BytesIO(data))
    src_fmt = im.format
    if getattr(im, "n_frames", 1) > 1:
        try:
            im.seek(0)
        except Exception:
            pass
    if im.mode == "P" and "transparency" in im.info:
        im = im.convert("RGBA")
    try:
        im = ImageOps.exif_transpose(im)
    except Exception:
        pass
    im.load()
    try:
        im.format = src_fmt          # exif_transpose 返回的是副本，会丢掉 format
    except Exception:
        pass
    return im


def _flatten(im: Image.Image, bg=(255, 255, 255)) -> Image.Image:
    """把带透明通道的图铺到白底上（JPEG 用）。"""
    if im.mode == "RGB":
        return im
    if im.mode in ("RGBA", "LA", "PA", "P", "La"):
        im2 = im.convert("RGBA")
        base = Image.new("RGB", im2.size, bg)
        base.paste(im2, mask=im2.split()[-1])
        return base
    return im.convert("RGB")


def _resize_max(im: Image.Image, max_dim: int) -> Image.Image:
    if not max_dim:
        return im
    w, h = im.size
    m = max(w, h)
    if m <= max_dim:
        return im
    k = max_dim / float(m)
    return im.resize((max(1, round(w * k)), max(1, round(h * k))), _LANCZOS)


def _resize_scale(im: Image.Image, k: float) -> Image.Image:
    if k >= 1.0:
        return im
    return im.resize((max(1, round(im.width * k)), max(1, round(im.height * k))), _LANCZOS)


def _pick_format(im_src: Image.Image, s: Settings, src_fmt: str) -> str:
    """决定输出格式。"""
    want = (s.img_format or "auto").lower()
    src_ext = {"PNG": "png", "JPEG": "jpeg", "MPO": "jpeg", "WEBP": "webp",
               "BMP": "png", "GIF": "png", "TIFF": "png", "AVIF": "avif",
               "JPEG2000": "png", "ICO": "png"}.get(src_fmt.upper(), "png")
    if want == "keep":
        return src_ext
    if want != "auto":
        return want
    if image_has_alpha(im_src):
        return "webp" if _webp_ok() else "png"
    return "jpeg"


def _looks_like_graphics(im: Image.Image, max_colors: int = 6000) -> bool:
    """判断是否像截图/图形（颜色数少、大片纯色），照片则返回 False。"""
    try:
        thumb = im if im.mode == "RGB" else im.convert("RGB")
        if max(thumb.size) > 220:
            thumb = thumb.copy()
            thumb.thumbnail((220, 220), _LANCZOS)
        colors = thumb.getcolors(maxcolors=300000)
        return colors is not None and len(colors) <= max_colors
    except Exception:
        return False


def _webp_ok() -> bool:
    try:
        from PIL import features
        return bool(features.check("webp"))
    except Exception:
        return False


def _avif_ok() -> bool:
    try:
        from PIL import features
        return bool(features.check("avif"))
    except Exception:
        return False


def encode_image(im: Image.Image, fmt: str, quality: int, s: Settings,
                 icc=None, exif=None) -> bytes:
    """按指定格式/质量编码成字节。"""
    fmt = fmt.lower()
    buf = io.BytesIO()
    kw = {}
    if s.keep_meta:
        if icc:
            kw["icc_profile"] = icc
        if exif and fmt in ("jpeg", "webp", "png"):
            kw["exif"] = exif
    if fmt == "jpeg":
        im2 = _flatten(im)
        im2.save(buf, "JPEG", quality=int(quality), optimize=True, progressive=True,
                 subsampling=0 if quality >= 85 else 2, **kw)
        return buf.getvalue()
    if fmt == "webp":
        q = min(100, int(quality) + 6)
        im2 = im if im.mode in ("RGB", "RGBA") else im.convert("RGBA" if image_has_alpha(im) else "RGB")
        im2.save(buf, "WEBP", quality=q, method=5, **kw)
        return buf.getvalue()
    if fmt == "avif":
        im2 = im if im.mode in ("RGB", "RGBA") else im.convert("RGB")
        im2.save(buf, "AVIF", quality=int(quality), speed=6)
        return buf.getvalue()
    if fmt == "bmp":
        _flatten(im).save(buf, "BMP")
        return buf.getvalue()
    # 默认 png（无损）
    im2 = im
    if im2.mode == "P":
        im2 = im2.convert("RGBA" if "transparency" in im2.info else "RGB")
    im2.save(buf, "PNG", optimize=True, compress_level=9)
    return buf.getvalue()


def _encode_webp_lossless(im: Image.Image, s: Settings, icc=None) -> bytes:
    buf = io.BytesIO()
    kw = {"icc_profile": icc} if (s.keep_meta and icc) else {}
    im2 = im if im.mode in ("RGB", "RGBA") else im.convert("RGBA")
    im2.save(buf, "WEBP", lossless=True, quality=100, method=5, **kw)
    return buf.getvalue()


def _fit_to_target_size(im: Image.Image, fmt: str, s: Settings, target: int,
                        icc=None, exif=None):
    """在质量/尺寸两个维度上搜索，尽量贴住目标体积。"""
    best = None                                   # (data, quality, scale, hit)
    lo, hi = 22, 96
    while lo <= hi:
        mid = (lo + hi) // 2
        data = encode_image(im, fmt, mid, s, icc, exif)
        if len(data) <= target:
            if best is None or len(data) > len(best[0]):
                best = (data, mid, 1.0, True)
            lo = mid + 1
        else:
            hi = mid - 1
    if best:
        return best[0], {"quality": best[1], "scale": best[2], "hit": True}
    for k in (0.85, 0.72, 0.60, 0.50, 0.42, 0.35, 0.28, 0.22, 0.17, 0.13):
        small = _resize_scale(im, k)
        for q in (80, 70, 60, 50, 40, 32, 25):
            data = encode_image(small, fmt, q, s, icc, exif)
            if len(data) <= target:
                return data, {"quality": q, "scale": k, "hit": True}
            if best is None or len(data) < len(best[0]):
                best = (data, q, k, False)
    if best:
        return best[0], {"quality": best[1], "scale": best[2], "hit": False}
    return None, {}


def compress_image(data: bytes, s: Settings, name: str = "") -> Result:
    """压缩一张图片，返回 Result。"""
    r = Result(kind="image")
    try:
        im_src = load_image(data)
    except Exception as exc:
        r.error = f"无法识别这个图片：{exc}"
        return r
    src_fmt = (im_src.format or "").upper()
    orig = len(data)
    icc = im_src.info.get("icc_profile")
    exif = im_src.info.get("exif")
    try:                                       # JPEG 大图先做 draft 抽样，解码快很多
        if src_fmt in ("JPEG", "MPO") and s.img_max_dim and max(im_src.size) > s.img_max_dim * 1.6:
            im_src.draft("RGB", (s.img_max_dim * 2, s.img_max_dim * 2))
    except Exception:
        pass
    im = _resize_max(im_src, s.img_max_dim)
    if s.img_gray:
        im = im.convert("L")
    fmt = _pick_format(im_src, s, src_fmt)

    def _keep_original(note: str) -> Result:
        r.ok = True
        r.data = data
        r.ext = Path(name).suffix.lstrip(".").lower() or (src_fmt.lower() or fmt)
        r.note = note
        r.stats = {"in": orig, "out": orig, "w0": im_src.width, "h0": im_src.height,
                   "w1": im_src.width, "h1": im_src.height, "fmt": (src_fmt or fmt).lower(),
                   "mode": "keep"}
        return r

    try:
        # ---- 无损档：PNG/WebP 无损重编码，JPEG 保留原文件 ----
        if s.preset == "lossless":
            if src_fmt in ("JPEG", "MPO"):
                return _keep_original("无损档：JPEG 无法无损重编码，已保留原文件（想变小请选有损档位）")
            cands = []
            cands.append(("png", encode_image(im, "png", 100, s)))
            if _webp_ok() and _looks_like_graphics(im):
                cands.append(("webp", _encode_webp_lossless(im, s, icc)))
            if cands:
                fmt2, best = min(cands, key=lambda kv: len(kv[1]))
                if len(best) < orig:
                    r.ok, r.data, r.ext = True, best, fmt2
                    r.stats = {"in": orig, "out": len(best),
                               "w0": im_src.width, "h0": im_src.height,
                               "w1": im.width, "h1": im.height, "fmt": fmt2,
                               "mode": "lossless"}
                    return r
            return _keep_original("无损档：JPEG 无法无损重编码，已保留原文件"
                                  if src_fmt in ("JPEG", "MPO") else "无损档：再压也不会更小，已保留原文件")

        # ---- 目标体积模式 ----
        if s.target_mode and s.target_kb > 0:
            if s.keep_if_bigger and orig <= s.target_kb * 1024:
                return _keep_original(f"原图 {human_size(orig)} 已经小于目标 {s.target_kb} KB，保持原文件")
            data2, meta = _fit_to_target_size(im, fmt, s, int(s.target_kb * 1024), icc, exif)
            if not data2:
                r.error = "目标体积太小，无法达标"
                return r
            if not meta.get("hit"):
                r.warnings.append("已尽力压缩，仍未达到目标体积")
            if s.keep_if_bigger and len(data2) >= orig:
                return _keep_original(f"再压也不会更小（目标 {s.target_kb} KB > 原图 {human_size(orig)}），保持原文件")
            w1, h1 = _resize_scale(im, meta.get("scale", 1.0)).size
            r.ok, r.data, r.ext = True, data2, fmt
            r.stats = {"in": orig, "out": len(data2), "w0": im_src.width, "h0": im_src.height,
                       "w1": w1, "h1": h1, "fmt": fmt,
                       "quality": meta.get("quality"), "mode": "target"}
            return r

        # ---- 常规质量档 ----
        data2 = encode_image(im, fmt, s.img_quality, s, icc, exif)
        # 截图/图形类（PNG 等无损来源）如果 WebP 无损明显更小，就换成它：画质更好、体积更小
        if (fmt == "jpeg" and s.img_format == "auto" and not s.img_gray
                and src_fmt in ("PNG", "BMP", "GIF", "TIFF", "PPM", "PGM", "TGA", "DDS")
                and _webp_ok() and not image_has_alpha(im_src)
                and _looks_like_graphics(im)):
            try:
                alt = _encode_webp_lossless(im, s, icc)
                if len(alt) * 2 < len(data2):
                    data2, fmt = alt, "webp"
                    r.note = "自动选用了无损 WebP：比 JPEG 小一半以上，而且没有压缩痕迹"
            except Exception:
                pass
        if s.keep_if_bigger and len(data2) >= orig:
            alt = None
            if fmt != "png":
                alt = encode_image(im, "png", 100, s)
            if alt is not None and len(alt) < orig:
                data2, fmt = alt, "png"
            else:
                return _keep_original(f"再压也不会更小（估计 {human_size(len(data2))}），已保留原图")
        r.ok, r.data, r.ext = True, data2, fmt
        r.stats = {"in": orig, "out": len(data2), "w0": im_src.width, "h0": im_src.height,
                   "w1": im.width, "h1": im.height, "fmt": fmt, "quality": s.img_quality,
                   "mode": "quality"}
        return r
    except Exception as exc:
        r.error = f"压缩失败：{exc}"
        r.warnings.append(traceback.format_exc(limit=3))
        return r

# ---------------------------------------------------------------- PDF 引擎
def pdf_page_count(data: bytes) -> int:
    if HAS_PIKEPDF:
        try:
            with pikepdf.open(io.BytesIO(data), suppress_warnings=True, attempt_recovery=True) as pdf:
                return len(pdf.pages)
        except Exception:
            pass
    if HAS_PDFIUM:
        try:
            doc = pdfium.PdfDocument(data)
            try:
                return len(doc)
            finally:
                doc.close()
        except Exception:
            pass
    return 0


def _walk_resources(res, out, seen, depth=0):
    """递归收集资源里的图片对象（含 Form XObject 里的）。"""
    if depth > 8 or not isinstance(res, pikepdf.Dictionary):
        return
    xo = res.get("/XObject")
    if xo is None:
        return
    try:
        items = list(xo.items())
    except Exception:
        return
    for _key, val in items:
        try:
            if not isinstance(val, pikepdf.Stream):
                continue
            sub = val.get("/Subtype")
        except Exception:
            continue
        if sub == "/Image":
            if val.objgen not in seen:
                seen.add(val.objgen)
                out.append(val)
        elif sub == "/Form":
            sub_res = val.get("/Resources")
            _walk_resources(sub_res if sub_res is not None else res, out, seen, depth + 1)


def _walk_annots(page, out, seen):
    """批注外观流里也可能藏着图片。"""
    try:
        annots = page.get("/Annots")
    except Exception:
        return
    if annots is None:
        return
    try:
        for a in annots:
            ap = a.get("/AP")
            if ap is None:
                continue
            for key in ("/N", "/R", "/D"):
                st = ap.get(key)
                if isinstance(st, pikepdf.Stream):
                    res = st.get("/Resources")
                    if res is not None:
                        _walk_resources(res, out, seen, 1)
    except Exception:
        return


def _collect_pdf_images(pdf):
    """收集整个文档里唯一的图片对象。"""
    seen, out = set(), []
    try:
        pages = list(pdf.pages)
    except Exception:
        return out
    for page in pages:
        res = None
        try:
            res = page.get("/Resources")
            if res is None:                              # 顺着页面树往上找继承的 /Resources
                parent = page.get("/Parent")
                depth = 0
                while res is None and parent is not None and depth < 16:
                    res = parent.get("/Resources")
                    parent = parent.get("/Parent")
                    depth += 1
        except Exception:
            res = None
        _walk_resources(res, out, seen, 0)
        _walk_annots(page, out, seen)
    return out


def _reencode_pdf_image(obj, s: Settings):
    """把一张内嵌图片重采样/重编码为 JPEG。返回 None 表示跳过。"""
    if s.pdf_quality <= 0:
        return None
    try:
        img = pikepdf.PdfImage(obj)
    except Exception:
        return None
    try:
        if img.image_mask or img.bits_per_component != 8:
            return None
        if "/SMask" in obj or "/Mask" in obj or "/Decode" in obj or "/DecodeParms" in obj:
            return None                                   # 有透明蒙版/特殊解码的，跳过更安全
        w, h = img.width, img.height
        if not w or not h or max(w, h) < 220:
            return None
        old_len = len(obj.read_raw_bytes())
    except Exception:
        return None
    if old_len < 25000:                                   # 小图省不下多少，风险不值当
        return None
    try:
        pil = img.as_pil_image()
    except Exception:
        return None
    try:
        pw, ph = pil.size
        scale = 1.0
        cap = s.pdf_img_max
        if cap and max(pw, ph) > cap:
            scale = cap / float(max(pw, ph))
        if scale < 1.0:
            pil = pil.resize((max(1, round(pw * scale)), max(1, round(ph * scale))), _LANCZOS)
        if s.pdf_gray:
            pil = pil if pil.mode == "L" else pil.convert("L")
            cs = "/DeviceGray"
        else:
            if pil.mode != "RGB":
                pil = _flatten(pil) if image_has_alpha(pil) else pil.convert("RGB")
            cs = "/DeviceRGB"
        buf = io.BytesIO()
        pil.save(buf, "JPEG", quality=int(s.pdf_quality), optimize=True,
                 progressive=True, subsampling=0 if s.pdf_quality >= 85 else 2)
        jpeg = buf.getvalue()
    except Exception:
        return None
    if len(jpeg) >= old_len * 0.93:                        # 不划算就保持原样
        return None
    try:
        obj.write(jpeg, filter=pikepdf.Name("/DCTDecode"))
        obj.Width = pil.width
        obj.Height = pil.height
        obj.BitsPerComponent = 8
        obj.ColorSpace = pikepdf.Name(cs)
        return {"old": old_len, "new": len(jpeg), "w": pil.width, "h": pil.height}
    except Exception:
        return None


def verify_pdf(data: bytes, expect_pages: int = 0):
    """输出自检：能用引擎打开、页数一致、首页能渲染出来。"""
    if HAS_PIKEPDF:
        try:
            with pikepdf.open(io.BytesIO(data), suppress_warnings=True, attempt_recovery=True) as pdf:
                n = len(pdf.pages)
                if n <= 0:
                    return False, "页数为 0"
                if expect_pages and n != expect_pages:
                    return False, f"页数不一致（{n} != {expect_pages}）"
        except Exception as exc:
            return False, f"输出无法打开：{exc}"
    if HAS_PDFIUM:
        try:
            doc = pdfium.PdfDocument(data)
            try:
                if len(doc) <= 0:
                    return False, "页数为 0"
                pg = doc[0]
                try:
                    bmp = pg.render(scale=0.5)
                    try:
                        im = bmp.to_pil()
                        if im.size[0] < 2 or im.size[1] < 2:
                            return False, "首页渲染结果异常"
                    finally:
                        bmp.close()
                finally:
                    pg.close()
            finally:
                doc.close()
        except Exception as exc:
            return False, f"首页渲染失败：{exc}"
    return True, ""


def _pdf_rasterize(data: bytes, s: Settings, progress=None) -> Result:
    """整页栅格化：体积最小，但文字会变成图片。"""
    r = Result(kind="pdf")
    if not HAS_PDFIUM:
        r.error = "缺少 pypdfium2，无法使用「整页栅格化」模式（请运行 安装依赖.bat）"
        return r
    try:
        doc = pdfium.PdfDocument(data)
    except Exception as exc:
        r.error = f"无法打开 PDF：{exc}"
        return r
    imgs = []
    try:
        n = len(doc)
        if n <= 0:
            r.error = "这个 PDF 没有页面"
            return r
        scale = max(0.08, float(s.pdf_raster_dpi) / 72.0)
        for i in range(n):
            if progress:
                progress(f"正在栅格化第 {i + 1}/{n} 页…")
            pg = doc[i]
            try:
                bmp = pg.render(scale=scale)
                try:
                    im = bmp.to_pil()
                finally:
                    bmp.close()
            finally:
                pg.close()
            if im.mode != "RGB":
                im = im.convert("RGB")
            if s.pdf_gray:
                im = im.convert("L")
            imgs.append(im)
    except Exception as exc:
        r.error = f"栅格化失败：{exc}"
        return r
    finally:
        try:
            doc.close()
        except Exception:
            pass
    try:
        buf = io.BytesIO()
        imgs[0].save(buf, "PDF", save_all=True, append_images=imgs[1:],
                     quality=int(max(20, s.pdf_quality)), resolution=float(s.pdf_raster_dpi))
        new = buf.getvalue()
    except Exception as exc:
        r.error = f"写出 PDF 失败：{exc}"
        return r
    r.ok, r.data, r.ext = True, new, "pdf"
    r.note = "整页栅格化：体积最小，但文字已变成图片（不可选中/搜索）"
    r.stats = {"in": len(data), "out": len(new), "pages": len(imgs),
               "engine": "raster", "dpi": s.pdf_raster_dpi, "mode": "raster"}
    return r


def _pdf_via_pikepdf(data: bytes, s: Settings, progress=None) -> Result:
    """结构保留压缩：逐张内嵌图片下采样重编码，文字与矢量原样保留。"""
    r = Result(kind="pdf")
    def say(msg):
        if progress:
            progress(msg)
    if not HAS_PIKEPDF:
        r.error = "缺少 pikepdf，无法做结构保留压缩"
        return r
    say("正在分析 PDF 结构…")
    try:
        pdf = pikepdf.open(io.BytesIO(data), suppress_warnings=True, attempt_recovery=True)
    except Exception as exc:
        try:
            pikepdf.open(io.BytesIO(data), password="", suppress_warnings=True)
        except Exception:
            pass
        r.error = f"无法打开 PDF：{exc}"
        return r
    try:
        pages = len(pdf.pages)
        imgs = _collect_pdf_images(pdf)
        touched, gained, skipped = 0, 0, 0
        for i, obj in enumerate(imgs):
            say(f"正在压缩内嵌图片 {i + 1}/{len(imgs)}…")
            try:
                info = _reencode_pdf_image(obj, s)
            except Exception:
                info = None
            if info:
                touched += 1
                gained += info["old"] - info["new"]
            else:
                skipped += 1
        say("正在重新打包 PDF（对象流 + 流压缩）…")
        out = io.BytesIO()
        mode = pikepdf.ObjectStreamMode.generate if s.pdf_objstm else pikepdf.ObjectStreamMode.preserve
        try:
            pdf.save(out, compress_streams=True, recompress_flate=True,
                     object_stream_mode=mode, linearize=False,
                     min_version="1.5" if s.pdf_objstm else "")
        except Exception:
            out = io.BytesIO()
            pdf.save(out, compress_streams=True, recompress_flate=True)
        new = out.getvalue()
        ok, why = verify_pdf(new, pages)
        if not ok:
            say("输出自检未通过，改用保守方式重试…")
            out2 = io.BytesIO()
            try:
                pdf.save(out2, compress_streams=True)
                new2 = out2.getvalue()
                ok2, why2 = verify_pdf(new2, pages)
            except Exception as exc:
                ok2, why2, new2 = False, str(exc), b""
            if ok2:
                new, ok, why = new2, True, ""
            else:
                r.warnings.append(f"压缩输出自检失败（{why}），已返回原文件")
                r.ok, r.data, r.ext = True, data, "pdf"
                r.note = "压缩输出未通过校验，为安全起见保持原文件"
                r.stats = {"in": len(data), "out": len(data), "pages": pages,
                           "engine": "pikepdf", "mode": "keep", "images": len(imgs)}
                return r
        r.ok, r.data, r.ext = True, new, "pdf"
        r.stats = {"in": len(data), "out": len(new), "pages": pages,
                   "images": len(imgs), "images_done": touched, "images_skipped": skipped,
                   "engine": "pikepdf", "mode": "structure",
                   "image_gain": gained}
        return r
    except Exception as exc:
        r.error = f"压缩失败：{exc}"
        r.warnings.append(traceback.format_exc(limit=4))
        return r
    finally:
        try:
            pdf.close()
        except Exception:
            pass


def compress_pdf(data: bytes, s: Settings, progress=None, name: str = "") -> Result:
    """PDF 压缩总入口：按参数选引擎，并保证「不会更小就保持原文件」。

    栅格化档位下会把「结构保留」和「整页栅格化」两种结果都算出来，自动取更小的那个。
    """
    cands = []
    if HAS_PIKEPDF:
        cands.append(_pdf_via_pikepdf(data, s, progress))
    if s.pdf_raster or not HAS_PIKEPDF:
        if HAS_PDFIUM:
            cands.append(_pdf_rasterize(data, s, progress))
    if not cands:
        r = Result(kind="pdf")
        r.error = "缺少 PDF 引擎（需要 pikepdf 或 pypdfium2），请先运行 安装依赖.bat"
        return r
    oks = [c for c in cands if c.ok]
    if not oks:
        return cands[0]
    if len(oks) > 1:
        oks.sort(key=lambda c: len(c.data))
        for extra in oks[1:]:
            for w in extra.warnings:
                oks[0].warnings.append(w)
        if oks[0].note:
            oks[0].note += "；"
        oks[0].note += "两种引擎都比过了，这是更小的那个"
    r = oks[0]
    if r.ok and s.keep_if_bigger and len(r.data) >= len(data):
        r.data = data
        r.ext = "pdf"
        r.note = f"再压也不会更小（估计 {human_size(len(r.data))}），已保持原文件"
        r.stats = dict(r.stats, out=len(data), mode="keep")
    return r


# ---------------------------------------------------------------- PDF 预览渲染
class PdfPreview:
    """PDF 页面渲染 + 缓存（只在工作线程里用）。"""

    def __init__(self, max_docs: int = 3, max_pages: int = 60):
        self.docs = {}          # key -> pdfium.PdfDocument
        self.order = []         # key 的使用顺序
        self.pages = {}         # (key, index, width) -> PIL.Image
        self.page_order = []
        self.max_docs = max_docs
        self.max_pages = max_pages

    def _doc(self, key, data):
        if key in self.docs:
            return self.docs[key]
        doc = pdfium.PdfDocument(data)
        self.docs[key] = doc
        self.order.append(key)
        while len(self.order) > self.max_docs:
            old = self.order.pop(0)
            d = self.docs.pop(old, None)
            if d is not None:
                try:
                    d.close()
                except Exception:
                    pass
            for ck in [k for k in self.pages if k[0] == old]:
                self.pages.pop(ck, None)
        return doc

    def count(self, key, data) -> int:
        try:
            return len(self._doc(key, data))
        except Exception:
            return 0

    def render(self, key, data, index: int = 0, width_px: int = 900):
        if not HAS_PDFIUM:
            return None, 0, 0
        try:
            doc = self._doc(key, data)
            n = len(doc)
            if n <= 0:
                return None, 0, 0
            index = max(0, min(int(index), n - 1))
            ck = (key, index, int(width_px))
            if ck in self.pages:
                return self.pages[ck], index, n
            pg = doc[index]
            try:
                wpt, hpt = pg.get_size()
                scale = max(0.15, min(4.0, float(width_px) / max(1.0, float(wpt))))
                bmp = pg.render(scale=scale)
                try:
                    im = bmp.to_pil()
                finally:
                    bmp.close()
            finally:
                pg.close()
            if im.mode != "RGB":
                im = im.convert("RGB")
            self.pages[ck] = im
            self.page_order.append(ck)
            while len(self.page_order) > self.max_pages:
                old = self.page_order.pop(0)
                self.pages.pop(old, None)
            return im, index, n
        except Exception:
            return None, 0, 0

    def drop(self, key):
        """释放某个文档的缓存（删除文件时调用）。"""
        d = self.docs.pop(key, None)
        if d is not None:
            try:
                d.close()
            except Exception:
                pass
        if key in self.order:
            self.order.remove(key)
        for ck in [k for k in self.pages if k[0] == key]:
            self.pages.pop(ck, None)
        for ck in [k for k in self.page_order if k[0] == key]:
            self.page_order.remove(ck)

    def close(self):
        for d in self.docs.values():
            try:
                d.close()
            except Exception:
                pass
        self.docs.clear()
        self.order.clear()
        self.pages.clear()
        self.page_order.clear()

# ---------------------------------------------------------------- 拖拽支持
try:                                     # tkdnd（社区成熟的拖放扩展，和 Tk 原生配合）
    if os.environ.get("PC_NO_DND"):       # 排障用：强制关掉拖拽
        raise ImportError("PC_NO_DND")
    from tkinterdnd2 import TkinterDnD as _TkinterDnD, DND_FILES as _DND_FILES
    HAS_DND = True
except Exception:                        # 没装也不影响其它功能
    _TkinterDnD = None
    _DND_FILES = None
    HAS_DND = False

# 拖放方法来自 DnDWrapper；tkdnd 的加载用官方 require()，失败就自动降级
_DnDMixin = getattr(_TkinterDnD, "DnDWrapper", object) if HAS_DND else object


# ---------------------------------------------------------------- 图形界面
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import threading as _threading


@dataclass
class Item:
    uid: int
    name: str
    kind: str                      # image / pdf
    data: bytes
    src_path: Path | None = None
    result: "Result | None" = None
    dirty: bool = True
    token: int = 0
    page: int = 0
    pages: int = 0
    note: str = ""
    error: str = ""


def result_ext(result: Result, fallback: str = "bin") -> str:
    ext = (result.ext or "").lower()
    if ext == "jpeg":
        ext = "jpg"
    return ext or fallback


class Worker(_threading.Thread):
    """后台工作线程：同一个 key 的任务只保留最后一个（拖动滑块时不会堆积）。"""

    def __init__(self, outq: "queue.Queue"):
        super().__init__(daemon=True, name="pc-worker")
        self.outq = outq
        self.cv = _threading.Condition()
        self.tasks = {}
        self.order = []
        self.stopped = False

    def submit(self, key: str, fn) -> None:
        with self.cv:
            if key in self.tasks:
                try:
                    self.order.remove(key)
                except ValueError:
                    pass
            self.tasks[key] = fn
            self.order.append(key)
            self.cv.notify_all()

    def drop_prefix(self, prefix: str) -> None:
        with self.cv:
            for k in [k for k in self.order if k.startswith(prefix)]:
                self.order.remove(k)
                self.tasks.pop(k, None)

    def stop(self) -> None:
        with self.cv:
            self.stopped = True
            self.cv.notify_all()

    def run(self) -> None:                              # pragma: no cover
        while True:
            with self.cv:
                while not self.order and not self.stopped:
                    self.cv.wait(0.4)
                if self.stopped:
                    return
                key = self.order.pop(0)
                fn = self.tasks.pop(key, None)
            if fn is None:
                continue
            try:
                fn()
            except Exception:
                self.outq.put(("error", traceback.format_exc(limit=5)))


class App(tk.Tk, _DnDMixin):
    """主窗口。"""

    def __init__(self, preload=None, preview="after"):
        super().__init__()
        self.title(f"{APP_NAME} · {APP_SUBTITLE}  v{APP_VERSION}")
        self.geometry("1240x880")
        self.minsize(1000, 700)
        try:
            ttk.Style().theme_use("vista")
        except Exception:
            pass

        self.outq: "queue.Queue" = queue.Queue()
        self.worker = Worker(self.outq)
        self.worker.start()
        self.pdfpreview = PdfPreview()

        self.items: list[Item] = []
        self.uid_seq = 0
        self.settings = preset_settings("balanced")      # 与界面默认选中的「平衡」保持一致
        self.active: Item | None = None

        self._prev_pil = {}         # (uid, variant, page) -> PIL
        self._photo = {}            # (uid, variant, page, w, h) -> PhotoImage
        self._img_src = {}          # uid -> PIL（原图预览用）
        self._job_started = 0.0
        self._busy = False
        self._last_slider_task = 0.0
        self._spin_guard = False

        self.preview_mode = tk.StringVar(value=preview if preview in ("before", "after", "side") else "after")
        self.var_preset = tk.StringVar(value="balanced")
        self.var_target = tk.BooleanVar(value=False)
        self.var_target_kb = tk.StringVar(value="500")
        self.var_status = tk.StringVar(value="把图片或 PDF 复制到剪贴板后按 Ctrl+V；也可以点「选择文件」，或直接把文件拖进窗口。")
        self.var_size = tk.StringVar(value="")
        self.var_page = tk.StringVar(value="- / -")

        self._build_ui()
        self.after(70, self._pump)
        for seq in ("<Control-v>", "<Control-V>", "<Control-Shift-V>"):
            self.bind_all(seq, self._on_paste_event)
        self.bind_all("<Delete>", lambda e: self.remove_selected())
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        cleanup_old_temp()
        self._dnd_ok = False
        if HAS_DND:                       # 注册窗口为拖放目标（失败不影响其它功能）
            try:
                _TkinterDnD.require(self)          # 官方 API：给已有的 Tk 根启用 tkdnd
                self.drop_target_register(_DND_FILES)
                self.dnd_bind("<<Drop>>", self._on_dnd)
                self._dnd_ok = True
            except Exception as exc:
                self._dnd_ok = False
                self._dnd_error = repr(exc)
        self._dlog(f"拖拽支持 = {self._dnd_ok} (HAS_DND={HAS_DND})")
        if preload:
            self.after(150, lambda: self.load_paths(preload))

    # ------------------------------------------------------------ 界面搭建
    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        self._build_toolbar()
        paned = ttk.PanedWindow(self, orient="horizontal")
        paned.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 6))
        self._build_list(paned)
        self._build_preview(paned)
        self._build_options()
        self._build_status()
        self._build_actions()
        self._sync_scales()

    def _build_toolbar(self):
        bar = ttk.Frame(self)
        bar.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 6))
        ttk.Button(bar, text="粘贴剪贴板 (Ctrl+V)", command=self.paste_clipboard).pack(side="left")
        ttk.Button(bar, text="选择文件…", command=self.choose_files).pack(side="left", padx=6)
        ttk.Button(bar, text="移除所选", command=self.remove_selected).pack(side="left")
        ttk.Button(bar, text="清空", command=self.clear_all).pack(side="left", padx=6)

        eng = []
        eng.append("pikepdf ✓" if HAS_PIKEPDF else "pikepdf ✗")
        eng.append("pypdfium2 ✓" if HAS_PDFIUM else "pypdfium2 ✗")
        label = ttk.Label(bar, text="  引擎：" + "   ".join(eng), foreground="#666")
        label.pack(side="right")
        ttk.Label(bar, text="  文件可直接拖进窗口  ", foreground="#8a6d00").pack(side="right")

    def _build_list(self, parent):
        frame = ttk.LabelFrame(parent, text="待处理文件")
        parent.add(frame, weight=0)
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        cols = ("kind", "orig", "new", "save", "state")
        self.tree = ttk.Treeview(frame, columns=cols, show="tree headings", height=14)
        self.tree.heading("#0", text="名称")
        self.tree.heading("kind", text="类型")
        self.tree.heading("orig", text="原始")
        self.tree.heading("new", text="压缩后")
        self.tree.heading("save", text="节省")
        self.tree.heading("state", text="状态")
        self.tree.column("#0", width=160, anchor="w")
        self.tree.column("kind", width=44, anchor="center", stretch=False)
        self.tree.column("orig", width=68, anchor="e", stretch=False)
        self.tree.column("new", width=68, anchor="e", stretch=False)
        self.tree.column("save", width=52, anchor="e", stretch=False)
        self.tree.column("state", width=72, anchor="w", stretch=False)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vs = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        vs.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=vs.set)
        self.tree.bind("<<TreeviewSelect>>", self._on_select)
        self.tree.tag_configure("err", foreground="#b00020")
        self.tree.tag_configure("keep", foreground="#0a6b3d")
        self.tree.tag_configure("wait", foreground="#8a6d00")
    # ------------------------------------------------------------ 预览面板
    def _build_preview(self, parent):
        frame = ttk.LabelFrame(parent, text="预览")
        parent.add(frame, weight=1)
        frame.rowconfigure(1, weight=1)
        frame.columnconfigure(0, weight=1)
        top = ttk.Frame(frame)
        top.grid(row=0, column=0, sticky="ew", padx=6, pady=(4, 2))
        for text, val in (("原图", "before"), ("压缩后", "after"), ("并排对比", "side")):
            ttk.Radiobutton(top, text=text, value=val, variable=self.preview_mode,
                            style="Toolbutton", command=self._on_preview_mode).pack(side="left", padx=1)
        self.pdf_nav = ttk.Frame(top)
        ttk.Button(self.pdf_nav, text="<", width=3, command=lambda: self.turn_page(-1)).pack(side="left")
        ttk.Label(self.pdf_nav, textvariable=self.var_page, width=10, anchor="center").pack(side="left", padx=3)
        ttk.Button(self.pdf_nav, text=">", width=3, command=lambda: self.turn_page(1)).pack(side="left")

        self.canvas = tk.Canvas(frame, bg="#202124", highlightthickness=0, height=380)
        self.canvas.grid(row=1, column=0, sticky="nsew", padx=6, pady=4)
        self.canvas.bind("<Configure>", self._on_canvas_resize)

        self.lbl_info = ttk.Label(frame, text="", anchor="w", justify="left",
                                  wraplength=880, foreground="#3b3b3b")
        self.lbl_info.grid(row=2, column=0, sticky="ew", padx=8, pady=(0, 6))
    # ------------------------------------------------------------ 设置面板
    def _build_options(self):
        outer = ttk.LabelFrame(self, text="压缩设置")
        outer.grid(row=2, column=0, sticky="ew", padx=10)
        row1 = ttk.Frame(outer)
        row1.pack(fill="x", padx=8, pady=(6, 2))
        ttk.Label(row1, text="档位：").pack(side="left")
        for key in ("lossless", "high", "balanced", "extreme"):
            ttk.Radiobutton(row1, text=PRESET_NAMES[key], value=key, variable=self.var_preset,
                            command=self.on_preset).pack(side="left", padx=(0, 3))
        ttk.Radiobutton(row1, text="自定义", value="custom", variable=self.var_preset,
                        command=self.on_preset).pack(side="left")
        ttk.Separator(row1, orient="vertical").pack(side="left", fill="y", padx=10)
        ttk.Checkbutton(row1, text="目标体积", variable=self.var_target,
                        command=self.on_target_toggle).pack(side="left")
        self.ent_target = ttk.Entry(row1, textvariable=self.var_target_kb, width=7)
        self.ent_target.pack(side="left", padx=4)
        ttk.Label(row1, text="KB（自动挑质量与尺寸）").pack(side="left")
        self.ent_target.bind("<Return>", lambda e: self.on_target_toggle())

        nb = ttk.Notebook(outer)
        nb.pack(fill="x", padx=8, pady=(2, 8))
        self.nb = nb
        self.tab_img = ttk.Frame(nb)
        self.tab_pdf = ttk.Frame(nb)
        nb.add(self.tab_img, text="图片设置")
        nb.add(self.tab_pdf, text="PDF 设置")
        self._scales = {}

        t = self.tab_img
        self._add_scale(t, 0, "最长边", "img_max_dim", 0, 6000, 20,
                        lambda v: "不缩放" if v == 0 else f"{v} px")
        self._add_scale(t, 1, "压缩质量", "img_quality", 30, 100, 1, lambda v: str(int(v)))
        ttk.Label(t, text="输出格式").grid(row=0, column=3, sticky="e", padx=(24, 4))
        self.cmb_fmt = ttk.Combobox(t, width=12, state="readonly", values=(
            "自动识别", "保持原格式", "JPEG", "WebP", "PNG", "AVIF"))
        self.cmb_fmt.grid(row=0, column=4, sticky="w")
        self.cmb_fmt.bind("<<ComboboxSelected>>", self._on_fmt)
        self.chk_img_gray = ttk.Checkbutton(t, text="转成灰度", command=self._on_toggle_img)
        self.chk_img_gray.grid(row=1, column=3, sticky="w", padx=(24, 4))
        self.chk_meta = ttk.Checkbutton(t, text="保留 EXIF / ICC 元数据", command=self._on_toggle_img)
        self.chk_meta.grid(row=1, column=4, sticky="w", pady=2)

        t = self.tab_pdf
        self._add_scale(t, 0, "内嵌图片最长边", "pdf_img_max", 0, 5000, 100,
                        lambda v: "不缩放" if v == 0 else f"{v} px")
        self._add_scale(t, 1, "图片质量", "pdf_quality", 0, 95, 1,
                        lambda v: "不重编码" if v == 0 else str(int(v)))
        self.chk_pdf_gray = ttk.Checkbutton(t, text="图片转灰度", command=self._on_toggle_pdf)
        self.chk_pdf_gray.grid(row=0, column=3, sticky="w", padx=(24, 4))
        self.chk_raster = ttk.Checkbutton(t, text="整页栅格化（更小，但文字变图片）",
                                          command=self._on_toggle_pdf)
        self.chk_raster.grid(row=1, column=3, columnspan=2, sticky="w", padx=(24, 4))
        self._add_scale(t, 2, "栅格化精度", "pdf_raster_dpi", 60, 300, 2, lambda v: f"{v} DPI")
        self.chk_objstm = ttk.Checkbutton(t, text="对象流重新打包", command=self._on_toggle_pdf)
        self.chk_objstm.grid(row=2, column=3, sticky="w", padx=(24, 4))
        ttk.Label(t, text="栅格化档位会自动和「保留文字」方案比大小，取更小的那个。",
                  foreground="#666").grid(row=3, column=0, columnspan=6, sticky="w", padx=10, pady=(4, 2))

    def _add_scale(self, parent, row, label, key, lo, hi, step, fmt):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="e", padx=(10, 6), pady=3)
        lbl = ttk.Label(parent, width=13, anchor="w")
        lbl.grid(row=row, column=2, sticky="w")

        def cb(v):
            val = int(round(float(v) / float(step)) * step)
            lbl.config(text=fmt(val))
            self._set_param(key, val)

        sc = ttk.Scale(parent, from_=lo, to=hi, orient="horizontal", length=220, command=cb)
        sc.grid(row=row, column=1, sticky="w")
        cur = getattr(self.settings, key)
        sc.set(cur)
        lbl.config(text=fmt(cur))
        self._scales[key] = (sc, lbl, fmt)
        return sc

    def _sync_scales(self):
        """把 settings 同步到控件（不触发重算）。"""
        self._spin_guard = True
        try:
            for key, (sc, lbl, fmt) in self._scales.items():
                val = getattr(self.settings, key)
                sc.set(val)
                lbl.config(text=fmt(val))
            pairs = ((self.chk_img_gray, self.settings.img_gray),
                     (self.chk_meta, self.settings.keep_meta),
                     (self.chk_pdf_gray, self.settings.pdf_gray),
                     (self.chk_raster, self.settings.pdf_raster),
                     (self.chk_objstm, self.settings.pdf_objstm))
            for widget, on in pairs:
                widget.state(["selected"] if on else ["!selected"])
            self.cmb_fmt.set({"auto": "自动识别", "keep": "保持原格式", "jpeg": "JPEG",
                              "webp": "WebP", "png": "PNG", "avif": "AVIF"}.get(
                                  self.settings.img_format, "自动识别"))
        finally:
            self._spin_guard = False

    # ------------------------------------------------------------ 状态栏 / 按钮
    def _build_status(self):
        bar = ttk.Frame(self)
        bar.grid(row=3, column=0, sticky="ew", padx=10, pady=(8, 2))
        self.lbl_size = ttk.Label(bar, textvariable=self.var_size, font=("Segoe UI", 11, "bold"))
        self.lbl_size.pack(side="left")
        self.pb = ttk.Progressbar(bar, mode="indeterminate", length=150)
        self.lbl_status = ttk.Label(bar, textvariable=self.var_status, foreground="#555")
        self.lbl_status.pack(side="left", padx=14)

    def _build_actions(self):
        bar = ttk.Frame(self)
        bar.grid(row=4, column=0, sticky="ew", padx=10, pady=(2, 10))
        ttk.Button(bar, text="导出全部…", command=self.export_all).pack(side="left")
        ttk.Button(bar, text="导出所选…", command=self.export_selected).pack(side="left", padx=6)
        ttk.Separator(bar, orient="vertical").pack(side="left", fill="y", padx=12)
        self.btn_copy = ttk.Button(bar, text="复制到剪贴板", command=lambda: self.copy_result(False))
        self.btn_copy.pack(side="left")
        self.btn_copy_file = ttk.Button(bar, text="复制为文件", command=lambda: self.copy_result(True))
        self.btn_copy_file.pack(side="left", padx=6)
        ttk.Button(bar, text="使用说明", command=self.show_help).pack(side="right")
    # ------------------------------------------------------------ 参数联动
    def _set_param(self, key, val):
        if self._spin_guard or getattr(self.settings, key) == val:
            return
        setattr(self.settings, key, val)
        self.settings.preset = "custom"
        self.var_preset.set("custom")
        self._schedule_recompress("active", 0.30)

    def _on_toggle_img(self):
        if self._spin_guard:
            return
        self.settings.img_gray = "selected" in self.chk_img_gray.state()
        self.settings.keep_meta = "selected" in self.chk_meta.state()
        self.settings.preset = "custom"
        self.var_preset.set("custom")
        self._schedule_recompress("active", 0.15)

    def _on_toggle_pdf(self):
        if self._spin_guard:
            return
        self.settings.pdf_gray = "selected" in self.chk_pdf_gray.state()
        self.settings.pdf_raster = "selected" in self.chk_raster.state()
        self.settings.pdf_objstm = "selected" in self.chk_objstm.state()
        self.settings.preset = "custom"
        self.var_preset.set("custom")
        self._schedule_recompress("active", 0.15)

    def _on_fmt(self, _event=None):
        if self._spin_guard:
            return
        mapping = {"自动识别": "auto", "保持原格式": "keep", "JPEG": "jpeg",
                   "WebP": "webp", "PNG": "png", "AVIF": "avif"}
        self.settings.img_format = mapping.get(self.cmb_fmt.get(), "auto")
        self.settings.preset = "custom"
        self.var_preset.set("custom")
        self._schedule_recompress("active", 0.15)

    def on_preset(self):
        key = self.var_preset.get()
        if key == "custom":
            self.settings.preset = "custom"
            return
        self.settings = preset_settings(key, self.settings)
        self.settings.preset = key
        self._sync_scales()
        for it in self.items:
            it.dirty = True
            it.token += 1
            self._update_row(it)
        self.var_status.set(f"已套用「{PRESET_NAMES.get(key, key)}」档位，正在估算…")
        self._schedule_recompress("all", 0.05)

    def on_target_toggle(self):
        self.settings.target_mode = bool(self.var_target.get())
        try:
            kb = max(20, int(float(self.var_target_kb.get())))
        except Exception:
            kb = 500
        self.var_target_kb.set(str(kb))
        self.settings.target_kb = kb
        self.var_status.set("目标体积模式：自动搜索质量与尺寸，尽量贴住目标大小" if self.settings.target_mode
                            else "已关闭目标体积模式")
        self._schedule_recompress("active", 0.12)

    def _on_preview_mode(self):
        self._request_preview()

    def _schedule_recompress(self, scope="active", delay=0.25):
        if getattr(self, "_after_job", None):
            try:
                self.after_cancel(self._after_job)
            except Exception:
                pass
        self._after_job = self.after(int(delay * 1000), lambda: self.recompress(scope))

    def recompress(self, scope="active"):
        self._after_job = None
        targets = [self.active] if scope == "active" else list(self.items)
        for it in targets:
            if it is not None:
                self._submit_compress(it)

    def _submit_compress(self, it: Item):
        it.token += 1
        token = it.token
        it.dirty = True
        it.error = ""
        s = self.settings.cloned()
        self._set_row_state(it, "计算中…")
        if self.active is it:
            self._busy = True
            self._start_busy()

        def fn():
            t0 = time.time()
            try:
                if it.kind == "pdf":
                    r = compress_pdf(it.data, s, progress=self._progress_cb(it.uid), name=it.name)
                else:
                    r = compress_image(it.data, s, name=it.name)
            except Exception as exc:
                r = Result(kind=it.kind, error=str(exc))
                r.warnings.append(traceback.format_exc(limit=4))
            self.outq.put(("compressed", it.uid, token, r, time.time() - t0))

        self.worker.submit(f"cmp:{it.uid}", fn)

    def _progress_cb(self, uid):
        last = [0.0]

        def cb(text):
            now = time.time()
            if now - last[0] > 0.12:
                last[0] = now
                self.outq.put(("status", text))
        return cb

    # ------------------------------------------------------------ 导入
    def _on_paste_event(self, _event=None):
        self.paste_clipboard()
        return "break"

    def paste_clipboard(self):
        try:
            kind, payload = clipboard_get()
        except Exception as exc:
            self.var_status.set(f"读取剪贴板失败：{exc}")
            return
        if kind == "image":
            buf = io.BytesIO()
            im = payload
            if im.mode not in ("RGB", "RGBA", "L"):
                im = im.convert("RGB")
            im.save(buf, "PNG")
            name = f"剪贴板截图_{time.strftime('%m%d_%H%M%S')}.png"
            self._add_item(name, "image", buf.getvalue(), None)
            return
        if kind == "files":
            added, skipped = 0, []
            for p in payload:
                k = kind_of(p.name)
                if not k:
                    skipped.append(p.name)
                    continue
                try:
                    data = p.read_bytes()
                except Exception as exc:
                    skipped.append(f"{p.name}（读取失败：{exc}）")
                    continue
                if data:
                    self._add_item(p.name, k, data, p)
                    added += 1
            msg = f"已导入 {added} 个文件"
            if skipped:
                msg += f"，跳过 {len(skipped)} 个不支持的内容：" + "、".join(skipped[:3])
            self.var_status.set(msg)
            return
        self.var_status.set("剪贴板里没有图片或文件。请先截图/复制文件，或者点「选择文件」。")

    def choose_files(self):
        exts = sorted(IMAGE_EXTS | PDF_EXTS)
        files = filedialog.askopenfilenames(
            title="选择图片或 PDF（可多选）",
            filetypes=[("图片和 PDF", " ".join("*" + e for e in exts)), ("所有文件", "*.*")])
        n = 0
        for f in files:
            p = Path(f)
            k = kind_of(p.name)
            try:
                data = p.read_bytes()
            except Exception as exc:
                self.var_status.set(f"读取失败：{p.name}（{exc}）")
                continue
            if not k:
                k = "pdf" if p.suffix.lower() == ".pdf" else "image"
            if data:
                self._add_item(p.name, k, data, p)
                n += 1
        if n:
            self.var_status.set(f"已导入 {n} 个文件")

    def _add_item(self, name, kind, data, src_path):
        self.uid_seq += 1
        it = Item(uid=self.uid_seq, name=name, kind=kind, data=data, src_path=src_path)
        self.items.append(it)
        self.tree.insert("", "end", iid=str(it.uid), text=name,
                         values=(("图片" if kind == "image" else "PDF"),
                                 human_size(len(data)), "…", "", "排队中"))
        self.tree.selection_set(str(it.uid))
        if kind == "pdf":
            self.nb.select(self.tab_pdf)
        else:
            self.nb.select(self.tab_img)
        self.active = it
        self._submit_compress(it)
        self._refresh_copy_label()
        return it

    def _item(self, iid) -> "Item | None":
        try:
            uid = int(iid)
        except Exception:
            return None
        for it in self.items:
            if it.uid == uid:
                return it
        return None

    def remove_selected(self):
        sel = list(self.tree.selection())
        if not sel:
            return
        for iid in sel:
            it = self._item(iid)
            if it is None:
                continue
            self.worker.drop_prefix(f"cmp:{it.uid}")
            self.worker.drop_prefix(f"prv:{it.uid}")
            for pre in ("before", "after"):
                self.pdfpreview.drop(f"{it.uid}:{pre}")
            self.items.remove(it)
            self.tree.delete(iid)
        self._photo.clear()
        self._prev_pil.clear()
        self.active = None
        rest = self.tree.get_children()
        if rest:
            self.tree.selection_set(rest[0])
        else:
            self.active = None
            self.var_size.set("")
            self.lbl_info.config(text="")
            self.var_page.set("- / -")
            self._draw_canvas()
            self.var_status.set("列表已清空")

    def clear_all(self):
        self.tree.delete(*self.tree.get_children())
        self.items.clear()
        self.active = None
        self._photo.clear()
        self._prev_pil.clear()
        self.var_size.set("")
        self.lbl_info.config(text="")
        self.var_page.set("- / -")
        self._draw_canvas()
        self.var_status.set("已清空列表")
    # ------------------------------------------------------------ 列表与状态刷新
    def _set_row_state(self, it: Item, state: str):
        try:
            vals = list(self.tree.item(str(it.uid), "values"))
            if len(vals) < 5:
                vals = ["", "", "", "", ""]
            vals[4] = state
            self.tree.item(str(it.uid), values=vals)
        except Exception:
            pass

    def _update_row(self, it: Item):
        r = it.result
        if r is None or not r.ok:
            new_txt, save_txt, state, tag = "…", "", "排队中", ()
            if r is not None and r.error:
                state, tag = "失败", ("err",)
        else:
            new_txt = human_size(r.size)
            pct = (1 - r.size / len(it.data)) * 100 if it.data else 0
            save_txt = f"{pct:.0f}%" if pct >= 0 else f"+{abs(pct):.0f}%"
            if r.stats.get("mode") == "keep":
                state, tag = "保持原图", ("keep",)
            else:
                state, tag = "完成", ()
            if it.dirty:
                state, tag = "待重算", ("wait",)
        try:
            self.tree.item(str(it.uid), values=(("图片" if it.kind == "image" else "PDF"),
                                                human_size(len(it.data)), new_txt, save_txt, state),
                           tags=tag)
        except Exception:
            pass

    def _update_status_ui(self):
        it = self.active
        if it is None:
            return
        r = it.result
        inb = len(it.data)
        if r is not None and r.ok:
            outb = r.size
            pct = (1 - outb / inb) * 100 if inb else 0
            sign = "↓" if pct >= 0 else "↑"
            self.var_size.set(f"{human_size(inb)}  {sign}{abs(pct):.1f}%  {human_size(outb)}")
        else:
            self.var_size.set(f"{human_size(inb)}  →  计算中…")
        parts = []
        st = (r.stats if (r and r.ok) else {}) or {}
        if it.kind == "image":
            if st.get("w0"):
                parts.append(f"{st['w0']}×{st['h0']} → {st.get('w1')}×{st.get('h1')}")
            if st.get("fmt"):
                parts.append("格式 " + str(st["fmt"]).upper())
            if st.get("quality") and st.get("mode") == "quality":
                parts.append(f"质量 {st['quality']}")
        else:
            pages = st.get("pages") or it.pages
            parts.append(f"{pages or '?'} 页")
            if st.get("images"):
                parts.append(f"内嵌图片 {st.get('images_done', 0)}/{st['images']} 张已重压")
            if st.get("engine") == "raster":
                parts.append("整页栅格化（文字已变图片）")
            elif st.get("engine") == "pikepdf":
                parts.append("保留文字与矢量")
        if r is not None and r.error:
            parts.append("错误：" + r.error)
        elif r is not None and r.note:
            parts.append(r.note)
        if r is not None and r.warnings:
            parts.append("提示：" + str(r.warnings[0]).splitlines()[0])
        self.lbl_info.config(text=" ｜ ".join(parts))

    def _on_select(self, _event=None):
        sel = self.tree.selection()
        if not sel:
            return
        it = self._item(sel[0])
        if it is None or it is self.active:
            return
        self.active = it
        self._photo.clear()
        self._prev_pil.clear()
        self._update_nav(it)
        self._refresh_copy_label()
        self._update_row(it)
        self._update_status_ui()
        if it.dirty and it.result is None:
            self._submit_compress(it)
        self._request_preview(force=True)

    def _update_nav(self, it: Item):
        if it.kind == "pdf":
            if not self.pdf_nav.winfo_ismapped():
                self.pdf_nav.pack(side="left", padx=14)
            n = it.pages or (it.result.stats.get("pages") if it.result else 0) or 0
            self.var_page.set(f"{it.page + 1} / {n}" if n else "- / ?")
        else:
            self.pdf_nav.pack_forget()
            self.var_page.set("- / -")

    def turn_page(self, delta: int):
        it = self.active
        if it is None or it.kind != "pdf":
            return
        n = it.pages or (it.result.stats.get("pages") if it.result else 0) or 1
        it.page = max(0, min(n - 1, it.page + delta))
        self.var_page.set(f"{it.page + 1} / {n}")
        self._photo.clear()
        self._draw_canvas()
        self._request_preview(force=True)

    # ------------------------------------------------------------ 预览渲染
    def _on_canvas_resize(self, _event=None):
        if getattr(self, "_after_resize", None):
            try:
                self.after_cancel(self._after_resize)
            except Exception:
                pass
        self._after_resize = self.after(180, self._after_resize_done)

    def _after_resize_done(self):
        self._after_resize = None
        self._draw_canvas()
        self._request_preview()

    def _request_preview(self, force: bool = False):
        it = self.active
        if it is None or not HAS_IMAGETK:
            self._draw_canvas()
            return
        cw = max(240, self.canvas.winfo_width())
        ch = max(200, self.canvas.winfo_height())
        mode = self.preview_mode.get()
        variants = ("before", "after") if mode == "side" else (mode,)
        for variant in variants:
            if variant == "after" and not (it.result and it.result.ok):
                if it.dirty:
                    continue
                self._prev_pil[(it.uid, variant, it.page)] = None
                continue
            key = (it.uid, variant, it.page)
            if not force and key in self._prev_pil:
                continue
            box_w = int(cw / 2 - 16) if mode == "side" else int(cw - 16)
            box_h = int(ch - 12)
            self.worker.submit(f"prv:{it.uid}:{variant}:{it.page}",
                               self._preview_task(it, variant, it.page, box_w, box_h))
        self._draw_canvas()

    def _preview_task(self, it: Item, variant: str, page: int, box_w: int, box_h: int):
        uid, kind = it.uid, it.kind

        def fn():
            im = None
            pages = 0
            try:
                if variant == "before":
                    data = it.data
                else:
                    data = it.result.data if (it.result and it.result.ok) else None
                if data:
                    if kind == "pdf":
                        im, _idx, pages = self.pdfpreview.render(f"{uid}:{variant}", data, page, box_w)
                    else:
                        im = load_image(data)
                if im is not None:
                    if im.mode not in ("RGB", "L"):
                        im = im.convert("RGB")
                    im.thumbnail((max(40, box_w), max(40, box_h)), _LANCZOS)
            except Exception:
                im = None
            self.outq.put(("preview", uid, variant, page, im, pages))

        return fn

    def _draw_canvas(self):
        c = self.canvas
        c.delete("all")
        W, H = c.winfo_width(), c.winfo_height()
        it = self.active
        if it is None:
            c.create_text(W // 2, H // 2 - 12, fill="#9aa0a6", font=("Segoe UI", 12),
                          text="压一压 · Ctrl+V 粘贴截图 / 图片 / PDF")
            c.create_text(W // 2, H // 2 + 16, fill="#7f858b", font=("Segoe UI", 10),
                          text="也可以点「选择文件…」，或直接把文件拖进这个窗口")
            return
        if not HAS_IMAGETK:
            c.create_text(W // 2, H // 2, text="（当前环境缺少 ImageTk，无法显示预览）", fill="#9aa0a6")
            return
        mode = self.preview_mode.get()
        page = it.page

        def blit(variant, x0, y0, bw, bh, caption):
            pil = self._prev_pil.get((it.uid, variant, page))
            if pil is None:
                text = "正在计算…" if (it.dirty and variant == "after") else "（暂无内容）"
                c.create_text(x0 + bw // 2, y0 + bh // 2, text=text, fill="#8a8f95",
                              font=("Segoe UI", 11))
                return
            k = min(bw / float(pil.width), bh / float(pil.height))
            fw, fh = max(1, int(pil.width * k)), max(1, int(pil.height * k))
            ck = (it.uid, variant, page, fw, fh)
            ph = self._photo.get(ck)
            if ph is None:
                ph = ImageTk.PhotoImage(pil.resize((fw, fh), _LANCZOS) if (fw, fh) != pil.size else pil)
                if len(self._photo) > 12:
                    self._photo.clear()
                self._photo[ck] = ph
            c.create_image(x0 + max(0, (bw - fw) // 2), y0 + max(0, (bh - fh) // 2),
                           anchor="nw", image=ph)
            c.create_text(x0 + 10, y0 + 8, anchor="nw", text=caption,
                          fill="#f1f3f4", font=("Segoe UI", 10, "bold"))

        if mode == "side":
            half = (W - 14) // 2
            blit("before", 4, 4, half - 4, H - 8, "原图")
            blit("after", half + 10, 4, half - 4, H - 8, "压缩后")
            c.create_line(half + 5, 6, half + 5, H - 6, fill="#4a4d51")
        else:
            blit(mode, 4, 4, W - 8, H - 8, "原图" if mode == "before" else "压缩后")
    # ------------------------------------------------------------ 忙碌指示
    def _start_busy(self):
        try:
            if not self.pb.winfo_ismapped():
                self.pb.pack(side="right")
            self.pb.start(18)
        except Exception:
            pass

    def _stop_busy(self):
        try:
            self.pb.stop()
            self.pb.pack_forget()
        except Exception:
            pass

    # ------------------------------------------------------------ 消息泵
    def _pump(self):
        try:
            while True:
                msg = self.outq.get_nowait()
                tag = msg[0]
                if tag == "compressed":
                    _tag, uid, token, r, dt = msg
                    it = self._item(str(uid))
                    if it is None or token != it.token:
                        continue
                    it.result = r
                    it.dirty = False
                    if r.ok and r.stats.get("pages"):
                        it.pages = int(r.stats["pages"])
                    self._dlog(f"压缩完成 {it.name}: {len(it.data)} -> {r.size if r.ok else 'FAIL ' + r.error}")
                    self._update_row(it)
                    if it is self.active:
                        self._busy = False
                        self._stop_busy()
                        self._update_status_ui()
                        self._update_nav(it)
                        self._request_preview(force=True)
                        if r.error:
                            self.var_status.set("失败：" + r.error[:150])
                        elif r.warnings:
                            self.var_status.set("完成（%.1f 秒），%s" % (dt, str(r.warnings[0]).splitlines()[0][:120]))
                        else:
                            self.var_status.set("完成（%.1f 秒）%s" % (dt, ("：" + r.note) if r.note else ""))
                elif tag == "preview":
                    _tag, uid, variant, page, im, pages = msg
                    it = self._item(str(uid))
                    if it is None or it is not self.active:
                        continue
                    self._prev_pil[(uid, variant, page)] = im
                    if pages and pages != it.pages:
                        it.pages = pages
                        self.var_page.set(f"{it.page + 1} / {pages}")
                    self._draw_canvas()
                elif tag == "status":
                    self.var_status.set(str(msg[1])[:170])
                elif tag == "export_one":
                    _tag, uid, r = msg
                    it = self._item(str(uid))
                    if it is not None and r.ok:
                        it.result = r
                        it.dirty = False
                        if r.stats.get("pages"):
                            it.pages = int(r.stats["pages"])
                        self._update_row(it)
                        if it is self.active:
                            self._update_status_ui()
                            self._request_preview(force=True)
                elif tag == "export_done":
                    _tag, n, dest, fails = msg
                    self._stop_busy()
                    self.var_status.set(f"导出完成：{n} 个文件" + (f"，{len(fails)} 个失败" if fails else ""))
                    if fails:
                        messagebox.showwarning(APP_NAME, "有文件导出失败：\n" + "\n".join(fails[:5]))
                    elif n:
                        if messagebox.askyesno(APP_NAME, f"已导出 {n} 个文件到：\n{dest}\n\n现在打开这个文件夹吗？"):
                            try:
                                os.startfile(dest)
                            except Exception:
                                pass
                elif tag == "error":
                    self._stop_busy()
                    self.var_status.set("内部错误：" + str(msg[1]).splitlines()[-1][:150])
        except queue.Empty:
            pass
        self.after(70, self._pump)

    # ------------------------------------------------------------ 导出
    def export_all(self):
        self._export(list(self.items))

    def export_selected(self):
        sel = [self._item(iid) for iid in self.tree.selection()]
        sel = [x for x in sel if x is not None]
        if not sel:
            self.var_status.set("先在列表里选中要导出的文件")
            return
        self._export(sel)

    def _export(self, items):
        if not items:
            self.var_status.set("列表是空的：先 Ctrl+V 粘贴或点「选择文件」")
            return
        folder = filedialog.askdirectory(title="选择导出到的文件夹")
        if not folder:
            return
        dest = Path(folder)
        s = self.settings.cloned()
        jobs = [(it.uid, it.name, it.kind, it.data) for it in items]

        def fn():
            ok_n, fails = 0, []
            for uid, name, kind, data in jobs:
                try:
                    if kind == "pdf":
                        r = compress_pdf(data, s, name=name)
                    else:
                        r = compress_image(data, s, name=name)
                except Exception as exc:
                    r = Result(kind=kind, error=str(exc))
                self.outq.put(("export_one", uid, r))
                if not r.ok:
                    fails.append(f"{name}：{r.error}")
                    continue
                ext = result_ext(r, "pdf" if kind == "pdf" else "png")
                target = unique_path(dest / f"{Path(name).stem or 'output'}_compressed.{ext}")
                try:
                    target.write_bytes(r.data)
                    ok_n += 1
                    self.outq.put(("status", f"已导出 {target.name}（{human_size(len(r.data))}）"))
                except Exception as exc:
                    fails.append(f"{name}：写入失败 {exc}")
            self.outq.put(("export_done", ok_n, str(dest), fails))

        self._start_busy()
        self.var_status.set(f"正在导出 {len(jobs)} 个文件…")
        self.worker.submit("export:all", fn)

    # ------------------------------------------------------------ 复制到剪贴板
    def _copy_label(self):
        it = self.active
        if it is None:
            return "复制到剪贴板"
        return "复制到剪贴板（文件）" if it.kind == "pdf" else "复制到剪贴板（图片）"

    def _refresh_copy_label(self):
        try:
            self.btn_copy.config(text=self._copy_label())
        except Exception:
            pass

    def _flash_copy(self, text):
        """把按钮文字短暂改成「已复制」，给出肉眼可见的反馈。"""
        try:
            self.btn_copy.config(text=text)
            if getattr(self, "_flash_job", None):
                try:
                    self.after_cancel(self._flash_job)
                except Exception:
                    pass
            self._flash_job = self.after(2000, self._refresh_copy_label)
        except Exception:
            pass

    def copy_result(self, as_file=False):
        """把压缩结果放进剪贴板：图片=位图，PDF/文件=文件；写完读回自检。"""
        it = self.active
        if it is None:
            self.var_status.set("先在列表里选一个文件")
            return
        r = it.result
        if r is None or not r.ok:
            msg = "这个文件还没压缩完成，等状态变成「完成」再点复制"
            if r is not None and r.error:
                msg = f"这个文件压缩失败：{r.error}"
            self._dialog(APP_NAME, msg)
            return
        ext = result_ext(r, "pdf" if it.kind == "pdf" else "png")
        try:
            tmp = new_temp_file("." + ext, "clip_")
            tmp.write_bytes(r.data)
            if as_file or it.kind == "pdf":
                verified = clipboard_set_files([tmp], verify=True)
                if not verified:
                    self._copy_failed("把文件写进剪贴板没有成功（剪贴板不可用）", it)
                    return
                self._flash_copy("✔ 已复制文件")
                self._set_status(
                    "✔ 剪贴板里现在是「文件」：到 资源管理器 / 微信 / 邮件 窗口里按 Ctrl+V 才能粘出文件；"
                    "在 Word / 聊天输入框里按 Ctrl+V 不会有反应（那里只能粘图片或文字）", "ok")
            else:
                how, verified = clipboard_set_image(tmp)
                if not verified:                              # 写进去了但读不回（剪贴板被占用）
                    self._copy_failed(f"图片写入方式 {how} 后没能确认剪贴板内容", it)
                else:
                    self._flash_copy("✔ 已复制图片")
                    self._set_status(
                        f"✔ 剪贴板里现在是图片（{how}，透明区域填白）：可直接粘到 微信 / QQ / Word / 画图 / 浏览器",
                        "ok")
        except Exception as exc:
            self._copy_failed(str(exc), it)

    # ------------------------------------------------------------ 杂项
    def _dialog(self, title, text, ask=False):
        """弹窗（自动化测试时用 PC_NO_DIALOG=1 关掉，避免阻塞）。"""
        self._set_status(("✘ " if not ask else "") + text.splitlines()[0], "err")
        if os.environ.get("PC_NO_DIALOG"):
            return False
        try:
            if ask:
                return bool(messagebox.askyesno(title, text))
            messagebox.showwarning(title, text)
        except Exception:
            pass
        return False

    def _copy_failed(self, err, item):
        """复制失败：说明原因 + 提供「改导出文件」的兜底。"""
        text = (f"复制到剪贴板失败：\n{err}\n\n"
                "常见原因：\n"
                "· 有别的程序占着剪贴板（截图/录屏/剪贴板增强类工具、微信、输入法等）\n"
                "· 本程序权限不够（可右键 →「以管理员身份运行」再试）\n"
                "· 系统剪贴板服务异常（重启资源管理器或重新登录可恢复）\n\n"
                "要不要先「导出到文件夹」把结果存成文件？")
        if self._dialog(APP_NAME, text, ask=True):
            try:
                self._export([item])
            except Exception:
                pass

    def _set_status(self, text, kind="info"):
        """写状态栏并按结果上色：ok=绿、err=红、info=灰。"""
        try:
            self.var_status.set(text)
            self.lbl_status.config(foreground={"ok": "#0a6b3d", "err": "#b00020"}.get(kind, "#555"))
        except Exception:
            pass

    def _dlog(self, msg):
        """调试日志（PC_DEBUG_DROP=1 时写到 stderr，正式使用无输出）。"""
        if os.environ.get("PC_DEBUG_DROP"):
            try:
                sys.stderr.write(f"[{time.strftime('%H:%M:%S')}] {msg}\n")
                sys.stderr.flush()
            except Exception:
                pass

    def _on_dnd(self, event):
        """tkdnd 拖放回调：把拖进来的路径交给 on_drop。"""
        try:
            raw = event.data
            try:
                paths = list(self.tk.splitlist(raw))       # 正确处理带空格/花括号的路径
            except Exception:
                paths = [raw]
            paths = [str(x).strip() for x in paths if str(x).strip()]
            self._dlog(f"拖入 {len(paths)} 个路径")
            if paths:
                self.on_drop(paths)
        except Exception as exc:
            self.var_status.set(f"拖拽处理失败：{exc}")
        return getattr(event, "action", "copy") or "copy"

    def on_drop(self, paths):
        """处理拖进窗口的文件/文件夹。"""
        files = []
        for raw in paths:
            try:
                p = Path(raw)
                if p.is_dir():
                    for sub in sorted(p.rglob("*")):
                        try:
                            if sub.is_file() and kind_of(sub.name):
                                files.append(sub)
                        except Exception:
                            continue
                        if len(files) >= 300:
                            break
                elif p.is_file():
                    files.append(p)
            except Exception:
                continue
            if len(files) >= 300:
                break
        self._dlog(f"on_drop 展开后 {len(files)} 个文件")
        if not files:
            self.var_status.set("拖进来的内容里没有能处理的图片或 PDF")
            return
        added, skipped = 0, []
        for f in files:
            k = kind_of(f.name)
            if not k:
                skipped.append(f.name)
                continue
            try:
                data = f.read_bytes()
            except Exception as exc:
                skipped.append(f"{f.name}（{exc}）")
                continue
            if data:
                self._dlog(f"  + {f.name} ({k}, {len(data)} bytes)")
                self._add_item(f.name, k, data, f)
                added += 1
        self._dlog(f"on_drop 完成：加入 {added} 个")
        msg = f"拖入并加入 {added} 个文件"
        if skipped:
            msg += f"，跳过 {len(skipped)} 个不支持的内容"
        self.var_status.set(msg)

    def load_paths(self, paths):
        """从命令行/拖拽预载文件。"""
        n = 0
        for f in paths:
            p = Path(f)
            if not p.is_file():
                continue
            k = kind_of(p.name)
            if not k:
                continue
            try:
                data = p.read_bytes()
            except Exception as exc:
                self.var_status.set(f"读取失败：{p.name}（{exc}）")
                continue
            if data:
                self._add_item(p.name, k, data, p)
                n += 1
        if n:
            self.var_status.set(f"已载入 {n} 个文件")

    def auto_test(self, clip_path, mode="", clip_kind="files"):
        """自动化验收：把文件放进剪贴板 → 触发 Ctrl+V 绑定 → 报告结果（供开发自测）。"""
        report = {}
        try:
            if mode in ("before", "after", "side"):
                self.preview_mode.set(mode)
            if clip_kind == "image":
                clipboard_set_image(Path(clip_path))
            else:
                clipboard_set_files([Path(clip_path)])
            self.tree.event_generate("<Control-v>")
            report["paste_event"] = "sent"
        except Exception as exc:
            report["paste_event"] = f"error: {exc}"

        def finish():
            out = {"items": len(self.items), "details": []}
            for it in self.items:
                d = {"name": it.name, "kind": it.kind, "in": len(it.data)}
                if it.result:
                    d.update(ok=it.result.ok, out=it.result.size, ext=it.result.ext,
                             mode=it.result.stats.get("mode"), note=it.result.note,
                             error=it.result.error)
                d["preview_before"] = (it.uid, "before", 0) in self._prev_pil
                d["preview_after"] = (it.uid, "after", 0) in self._prev_pil
                d["preview_imgs"] = sum(1 for k in self._prev_pil if k[0] == it.uid and self._prev_pil[k])
                out["details"].append(d)
            out["status"] = self.var_status.get()
            out["size_label"] = self.var_size.get()
            out["info_label"] = self.lbl_info.cget("text")
            out["paste_event"] = report.get("paste_event")
            try:
                sys.stderr.write("AUTOTEST " + json.dumps(out, ensure_ascii=False) + "\n")
                sys.stderr.flush()
            except Exception:
                pass
            self.after(1200, self.destroy)

        self.after(7000, finish)

    def export_auto(self, folder):
        """自动化验收：导出全部（供开发自测）。"""
        dest = Path(folder)
        dest.mkdir(parents=True, exist_ok=True)
        s = self.settings.cloned()
        res = []
        for it in self.items:
            r = compress_pdf(it.data, s, name=it.name) if it.kind == "pdf" else compress_image(it.data, s, name=it.name)
            if r.ok:
                f = dest / f"{Path(it.name).stem}_compressed.{result_ext(r, 'bin')}"
                f.write_bytes(r.data)
                res.append([f.name, len(r.data)])
            else:
                res.append([it.name, "FAIL " + r.error])
        for it in self.items:
            self.copy_result(False)
            res.append(["clipboard", self.var_status.get()])
            break
        try:
            sys.stderr.write("EXPORTTEST " + json.dumps(res, ensure_ascii=False) + "\n")
            sys.stderr.flush()
        except Exception:
            pass
        self.after(500, self.destroy)

    def show_help(self):
        for md in (APP_DIR / "使用说明.md", BUNDLE_DIR / "使用说明.md"):
            try:
                if md.is_file():
                    os.startfile(md)
                    return
            except Exception:
                pass
        messagebox.showinfo(APP_NAME, HELP_TEXT)

    def _on_close(self):
        try:
            self.worker.stop()
        except Exception:
            pass
        try:
            self.pdfpreview.close()
        except Exception:
            pass
        self.destroy()


def make_app(**kwargs):
    """创建主窗口（拖放加载失败会在内部自动降级，不影响启动）。"""
    return App(**kwargs)


HELP_TEXT = (
    "【压一压】图片 / PDF 轻量压缩器\n"
    "粘贴一下，压一压，图片和 PDF 的体积就小了。离线运行、免安装、不覆盖原文件。\n\n"
    "【怎么用】\n"
    "1. 截图后直接 Ctrl+V；或在资源管理器里复制图片/PDF 后 Ctrl+V；\n"
    "   也可以点「选择文件…」，或者直接把文件（甚至整个文件夹）拖进窗口。\n"
    "2. 选档位：无损 / 高质量 / 平衡 / 极限；想精确控制就勾「目标体积」填 KB。\n"
    "3. 右边可以看原图与压缩后的对比；PDF 用 < > 翻页。\n"
    "4. 「导出全部…」存成新文件（不会覆盖原文件）；「复制到剪贴板」把结果直接拿去粘贴。\n\n"
    "【说明】\n"
    "· 图片：自动处理旋转、透明通道；无损档会优先尝试无损 WebP。\n"
    "· PDF：默认只重压内嵌图片，文字和矢量原样保留；「极限」档会额外尝试整页栅格化，取更小的结果。\n"
    "· 如果压缩后反而更大，程序会保留原文件并提示。\n"
    "· 全程本地处理，不联网、不上传。")

# ---------------------------------------------------------------- 命令行模式
def cli_main(argv):
    """python PasteCompress.py --cli <文件...> [--out=目录] [--preset=平衡] [--target-kb=300]"""
    pos, flags = [], {}
    for a in argv:
        if a.startswith("--"):
            if "=" in a:
                k, v = a[2:].split("=", 1)
                flags[k] = v
            else:
                flags[a[2:]] = True
        else:
            pos.append(a)
    pos = [p for p in pos if p != "cli"]
    if not pos:
        print("用法: python PasteCompress.py --cli <文件或文件夹...> [--out=目录] [--preset=balanced] [--target-kb=300]")
        return 2
    preset_key = str(flags.get("preset", "balanced"))
    if preset_key not in PRESET_NAMES:
        preset_key = {"无损": "lossless", "高质量": "high", "平衡": "balanced", "极限": "extreme"}.get(preset_key, "balanced")
    s = preset_settings(preset_key)
    if flags.get("target-kb"):
        s.target_mode = True
        s.target_kb = max(20, int(float(flags["target-kb"])))
    outdir = Path(str(flags["out"])).expanduser() if flags.get("out") and flags["out"] is not True else None
    if outdir:
        outdir.mkdir(parents=True, exist_ok=True)

    files = []
    for p in pos:
        pp = Path(p)
        if pp.is_dir():
            files += sorted(x for x in pp.iterdir() if x.is_file() and kind_of(x.name))
        elif pp.is_file():
            files.append(pp)
    if not files:
        print("没有找到可处理的文件")
        return 2

    total_in = total_out = 0
    bad = 0
    for p in files:
        kind = kind_of(p.name) or ("pdf" if p.suffix.lower() == ".pdf" else "image")
        try:
            data = p.read_bytes()
        except Exception as exc:
            print(f"[跳过] {p.name}: {exc}")
            bad += 1
            continue
        t0 = time.time()
        r = compress_pdf(data, s, name=p.name) if kind == "pdf" else compress_image(data, s, name=p.name)
        dt = time.time() - t0
        if not r.ok:
            print(f"[失败] {p.name}: {r.error}")
            bad += 1
            continue
        ext = result_ext(r, p.suffix.lstrip(".").lower())
        dest = unique_path((outdir or p.parent) / f"{p.stem}_compressed.{ext}")
        try:
            dest.write_bytes(r.data)
        except Exception as exc:
            print(f"[失败] {p.name}: 写入 {dest} 出错 {exc}")
            bad += 1
            continue
        total_in += len(data)
        total_out += len(r.data)
        pct = (1 - len(r.data) / len(data)) * 100 if data else 0
        print(f"[完成] {p.name}  {human_size(len(data))} -> {human_size(len(r.data))} "
              f"({pct:+.1f}%)  {dt:.1f}s  ->  {dest}")
        if r.note:
            print(f"        {r.note}")
    if total_in:
        print(f"\n合计：{human_size(total_in)} -> {human_size(total_out)}  节省 {(1 - total_out / total_in) * 100:.1f}%")
    return 1 if bad else 0


# ---------------------------------------------------------------- 自检
def _make_photo(w, h, seed=1):
    from PIL import ImageOps as _ops
    chans = [Image.effect_noise((w, h), 38 + 8 * i) for i in range(3)]
    im = Image.merge("RGB", chans)
    grad = _ops.colorize(Image.linear_gradient("L").resize((w, h)), (18, 38, 88), (252, 206, 128)).convert("RGB")
    return Image.blend(im, grad, 0.45)


def _make_screenshot(w=1400, h=900):
    from PIL import ImageDraw, ImageFont
    im = Image.new("RGB", (w, h), "white")
    d = ImageDraw.Draw(im)
    font = None
    for path in ("C:/Windows/Fonts/msyh.ttc", "C:/Windows/Fonts/simhei.ttf", "C:/Windows/Fonts/arial.ttf"):
        try:
            font = ImageFont.truetype(path, 22)
            break
        except Exception:
            font = None
    if font is None:
        font = ImageFont.load_default()
    for i in range(16):
        d.text((40, 40 + i * 38), f"第 {i + 1} 行：测试压缩质量 ABC abc 12345 中文示例文字。",
               fill=(25, 25, 35), font=font)
    for i in range(5):
        d.rectangle([40 + i * 250, 720, 250 + i * 250, 840], fill=(40 + i * 40, 110, 190 - i * 25))
    return im


def _make_flate_pdf(path, img):
    """造一个「整页 Flate 大图」的 PDF，用来测纯无损重打包与重编码。"""
    import zlib as _zlib
    img = img.convert("RGB")
    W, H = img.size
    comp = _zlib.compress(img.tobytes(), 6)
    cw = W // 2
    content = f"q {cw} 0 0 {cw * H // W} 40 60 cm /Im0 Do Q".encode()
    ccomp = _zlib.compress(content, 6)
    objs = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {cw + 80} {cw * H // W + 120}] "
        f"/Resources << /XObject << /Im0 4 0 R >> >> /Contents 5 0 R >>".encode(),
        f"<< /Type /XObject /Subtype /Image /Width {W} /Height {H} /ColorSpace /DeviceRGB "
        f"/BitsPerComponent 8 /Filter /FlateDecode /Length {len(comp)} >>".encode(),
        f"<< /Filter /FlateDecode /Length {len(ccomp)} >>".encode(),
    ]
    streams = {4: comp, 5: ccomp}
    out = bytearray(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
    offs = {}
    for i, o in enumerate(objs, start=1):
        offs[i] = len(out)
        out += f"{i} 0 obj\n".encode() + o + b"\n"
        if i in streams:
            out += b"stream\n" + streams[i] + b"\nendstream\n"
        out += b"endobj\n"
    xref = len(out)
    out += f"xref\n0 {len(objs) + 1}\n".encode() + b"0000000000 65535 f \n"
    for i in range(1, len(objs) + 1):
        out += f"{offs[i]:010d} 00000 n \n".encode()
    out += f"trailer\n<< /Size {len(objs) + 1} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF\n".encode()
    Path(path).write_bytes(bytes(out))


def run_selftest(with_clipboard=False):
    """自检：造样本 → 逐档压缩 → 校验体积/可打开性/像素差异。"""
    out = APP_DIR / "自检结果"
    shutil.rmtree(out, ignore_errors=True)
    out.mkdir(parents=True, exist_ok=True)
    print(f"{APP_NAME} v{APP_VERSION} 自检")
    print(f"Python {sys.version.split()[0]}  程序目录 {APP_DIR}")
    print(f"引擎：pikepdf={'有' if HAS_PIKEPDF else '无'}  pypdfium2={'有' if HAS_PDFIUM else '无'}  "
          f"Pillow={getattr(Image, '__version__', '?')}  WebP={_webp_ok()}  AVIF={_avif_ok()}  ImageTk={HAS_IMAGETK}")
    fails = []

    img_jpg = out / "样本_照片.jpg"
    _make_photo(1800, 1200, 2).save(img_jpg, "JPEG", quality=95, subsampling=0)
    img_png = out / "样本_截图.png"
    _make_screenshot().save(img_png)
    pdf_album = out / "样本_相册.pdf"
    pg = [_make_photo(1400, 1900, s) for s in (3, 4)]
    pg[0].save(pdf_album, save_all=True, append_images=pg[1:], quality=92, resolution=150.0)
    pdf_flate = out / "样本_整页图.pdf"
    _make_flate_pdf(pdf_flate, _make_screenshot(1200, 800))
    samples = [(img_jpg, "image"), (img_png, "image"), (pdf_album, "pdf"), (pdf_flate, "pdf")]

    print("\n--- 逐档压缩 ---")
    print(f"{'档位':<7}{'样本':<18}{'原始':>10}{'压缩后':>10}{'节省':>7}  校验")
    for preset in ("lossless", "high", "balanced", "extreme"):
        s = preset_settings(preset)
        for path, kind in samples:
            data = path.read_bytes()
            t0 = time.time()
            r = compress_pdf(data, s, name=path.name) if kind == "pdf" else compress_image(data, s, name=path.name)
            dt = time.time() - t0
            checks = []
            if not r.ok:
                checks.append("失败：" + r.error)
                fails.append(f"{preset}/{path.name}: {r.error}")
            else:
                if len(r.data) > len(data) and r.stats.get("mode") != "keep":
                    checks.append("体积变大")
                    fails.append(f"{preset}/{path.name}: 体积变大")
                if kind == "pdf":
                    ok2, why = verify_pdf(r.data, 0)
                    checks.append("PDF可打开" if ok2 else "打不开：" + why)
                    if not ok2:
                        fails.append(f"{preset}/{path.name}: {why}")
                else:
                    try:
                        im = Image.open(io.BytesIO(r.data))
                        im.load()
                        checks.append(f"{im.format} {im.size[0]}x{im.size[1]}")
                    except Exception as exc:
                        checks.append("解码失败")
                        fails.append(f"{preset}/{path.name}: 解码失败 {exc}")
                pct = (1 - len(r.data) / len(data)) * 100 if data else 0
                checks.insert(0, f"{pct:+.0f}%")
            print(f"{PRESET_NAMES[preset]:<7}{path.name:<18}{human_size(len(data)):>10}"
                  f"{human_size(len(r.data) if r.ok else 0):>10}{'':>7}  {' | '.join(checks)}  {dt:.1f}s")

    print("\n--- 目标体积模式（图片）---")
    s = preset_settings("balanced")
    s.target_mode = True
    for kb in (300, 100, 40):
        s.target_kb = kb
        for path, kind in samples[:2]:
            data = path.read_bytes()
            r = compress_image(data, s, name=path.name)
            if not r.ok:
                fails.append(f"目标 {kb}KB {path.name}: {r.error}")
                continue
            got = len(r.data) / 1024.0
            mark = "达标" if got <= kb * 1.05 else ("接近" if got <= kb * 1.4 else "未达标")
            if got > kb * 1.4:
                fails.append(f"目标 {kb}KB {path.name}: 得到 {got:.0f}KB")
            print(f"  目标 {kb:>4} KB  {path.name:<18}{human_size(len(data)):>10} -> {human_size(len(r.data)):>10}  {mark}")

    if with_clipboard:
        print("\n--- 剪贴板（会覆盖当前剪贴板内容）---")
        try:
            p = out / "剪贴板测试.png"
            _make_screenshot(320, 220).save(p)
            clipboard_set_image(p)
            kind, payload = clipboard_get()
            ok2 = kind == "image" and getattr(payload, "size", None) is not None
            print(f"  位图写入并读回：{'OK' if ok2 else '失败 ' + str(kind)}")
            if not ok2:
                fails.append("剪贴板位图测试失败")
            clipboard_set_files([p])
            kind2, payload2 = clipboard_get()
            ok3 = kind2 == "files" and payload2 and Path(payload2[0]).name == p.name
            print(f"  文件写入并读回：{'OK' if ok3 else '失败 ' + str(kind2)}")
            if not ok3:
                fails.append("剪贴板文件测试失败")
        except Exception as exc:
            print("  剪贴板测试异常：", exc)
            fails.append(f"剪贴板测试异常: {exc}")

    print("\n--- 结果 ---")
    if fails:
        print(f"发现 {len(fails)} 个问题：")
        for f in fails[:20]:
            print("  ·", f)
    else:
        print("全部通过 ✓")
    print(f"样本与压缩结果在：{out}")
    return 1 if fails else 0


def main():
    argv = sys.argv[1:]
    report_file = None
    for a in argv:
        if a.startswith("--report="):
            try:
                report_file = open(APP_DIR / a.split("=", 1)[1] if not os.path.isabs(a.split("=", 1)[1])
                                   else a.split("=", 1)[1], "w", encoding="utf-8")
                sys.stdout = report_file
                sys.stderr = report_file
            except Exception:
                report_file = None
    if "--selftest" in argv:
        return run_selftest(with_clipboard="--selftest-clip" in argv)
    if "--clip-test" in argv:
        return run_selftest(with_clipboard=True)
    if "--cli" in argv:
        return cli_main(argv)
    if "-h" in argv or "--help" in argv or "-?" in argv:
        print(__doc__)
        return 0
    preload = [a.split("=", 1)[1].strip('"') for a in argv if a.startswith("--load=")]
    for a in argv:
        if not a.startswith("--") and a.strip('"') and Path(a.strip('"')).is_file():
            preload.append(a.strip('"'))
    preview = "after"
    for a in argv:
        if a.startswith("--preview="):
            preview = a.split("=", 1)[1]
    autotest = [a.split("=", 1)[1] for a in argv if a.startswith("--autotest=")]
    autotest_mode = ""
    autotest_clip = "files"
    for a in argv:
        if a.startswith("--autotest-mode="):
            autotest_mode = a.split("=", 1)[1]
        elif a.startswith("--autotest-clip="):
            autotest_clip = a.split("=", 1)[1]
    export_to = [a.split("=", 1)[1] for a in argv if a.startswith("--export-test=")]
    if os.environ.get("PC_DEBUG_DUMP"):
        import faulthandler
        try:
            faulthandler.dump_traceback_later(10, repeat=True,
                                              file=open(APP_DIR / "stack.txt", "w", encoding="utf-8"))
        except Exception:
            pass
    try:
        app = make_app(preload=preload, preview=preview)
    except Exception:
        traceback.print_exc()
        if "--debug" not in argv:
            print("界面启动失败，请把上面的报错发给开发者；也可以先用命令行模式：\n"
                  "  python PasteCompress.py --cli 文件.pdf --preset=balanced")
        return 3
    if autotest:
        app.after(400, lambda: app.auto_test(autotest[0], autotest_mode, autotest_clip))
    if export_to and preload:
        app.after(2500, lambda: app.export_auto(export_to[0]))
    app.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main() or 0)
