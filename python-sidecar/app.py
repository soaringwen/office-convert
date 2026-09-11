"""
Office-Convert Python Sidecar
=============================

基于 MarkItDown + oletools 的兜底转换服务（可选组件）。

作用：
1. /convert      —— 调用 MarkItDown 将 Java 原生转换器未覆盖的格式
                   （.rtf、.odt/.ods/.odp、.xlsb 等）转换为 Markdown；
2. /inspect/ole  —— 使用 oletools (olefile) 枚举 OLE2 容器内部流结构，
                   用于排查旧版 Office 嵌入对象；
3. /vba          —— 使用 olevba 静态检测 VBA 宏（只读分析，绝不执行宏）。

安全约束（需求 10.1）：
- 不执行任何宏、脚本或嵌入程序；
- 不访问任何外部网络资源；
- 上传文件仅在临时目录中处理，请求结束后删除。

启动：uvicorn app:app --host 127.0.0.1 --port 8900
"""

import os
import shutil
import tempfile

from fastapi import FastAPI, File, HTTPException, UploadFile

app = FastAPI(title="office-convert sidecar", version="1.0.0")

MAX_SIZE = 20 * 1024 * 1024  # 与主服务一致：单文件 20MB

_markitdown = None


def get_markitdown():
    global _markitdown
    if _markitdown is None:
        from markitdown import MarkItDown
        _markitdown = MarkItDown()
    return _markitdown


@app.get("/health")
def health():
    return {"status": "ok", "markitdown": _markitdown is not None or _try_import()}


def _try_import():
    try:
        import markitdown  # noqa: F401
        return True
    except ImportError:
        return False


@app.post("/convert")
async def convert(file: UploadFile = File(...)):
    """转换为 Markdown（MarkItDown）。"""
    data = await file.read()
    if len(data) > MAX_SIZE:
        raise HTTPException(status_code=413, detail="文件超过 20MB 限制")

    tmp = tempfile.mkdtemp(prefix="officeconvert-")
    try:
        safe_name = os.path.basename(file.filename or "upload.bin") or "upload.bin"
        path = os.path.join(tmp, safe_name)
        with open(path, "wb") as f:
            f.write(data)
        result = get_markitdown().convert(path)
        markdown = getattr(result, "text_content", "") or ""
        return {"markdown": markdown, "filename": safe_name}
    except Exception as e:
        raise HTTPException(status_code=422, detail=f"转换失败: {e}")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


@app.post("/inspect/ole")
async def inspect_ole(file: UploadFile = File(...)):
    """枚举 OLE2 容器流结构（olefile，只读）。"""
    try:
        import olefile
    except ImportError:
        raise HTTPException(status_code=501, detail="oletools 未安装")

    data = await file.read()
    if len(data) > MAX_SIZE:
        raise HTTPException(status_code=413, detail="文件超过 20MB 限制")

    import io
    try:
        ole = olefile.OleFileIO(io.BytesIO(data))
    except Exception as e:
        raise HTTPException(status_code=422, detail=f"非 OLE2 容器: {e}")

    # olefile.listdir 返回路径元组列表
    streams = ["/".join(p) for p in ole.listdir(streams=True, storages=True)]
    meta = {}
    try:
        meta = ole.get_metadata().dump_dict() if hasattr(ole, "get_metadata") else {}
    except Exception:
        pass
    ole.close()
    return {"streams": streams, "metadata": meta}


@app.post("/vba")
async def vba_scan(file: UploadFile = File(...)):
    """VBA 宏静态检测（olevba，只分析不执行）。"""
    try:
        from oletools.olevba import VBA_Parser
    except ImportError:
        raise HTTPException(status_code=501, detail="oletools 未安装")

    data = await file.read()
    if len(data) > MAX_SIZE:
        raise HTTPException(status_code=413, detail="文件超过 20MB 限制")

    import io
    try:
        parser = VBA_Parser(io.BytesIO(data))
        found = parser.detect_vba_macros()
        modules = []
        if found:
            for (_, _, vba_filename, vba_code) in parser.extract_macros():
                modules.append({"module": vba_filename, "lines": len(vba_code.splitlines())})
        parser.close()
    except Exception as e:
        raise HTTPException(status_code=422, detail=f"分析失败: {e}")
    return {"has_macros": bool(found), "modules": modules,
            "note": "仅静态分析，绝不执行宏（需求 10.1）"}
