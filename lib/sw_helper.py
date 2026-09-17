"""SolidWorks 自动化公共模块：连接 + 参数化长方体建模示例。

用法:
    from sw_helper import connect
    sw, mtyp, ext, skmgr, featmgr = connect()   # 附着/启动 SW，返回类型化接口组
经验固化（本机 SW 2024 SP5, win32com 动态 dispatch）:
- SelectByID2 等带 VARIANT 参数的方法必须走 makepy 类型化包装（swnsis.py），否则报"类型不匹配"。
- RevisionNumber/GetTitle/GetType/GetDocumentCount/Volume 是属性形态；ClearSelection2(True)/EditRebuild3() 是方法。
- 无 .prtdot 模板 → styp.INewPart() 空白新建。
- CreateCenterRectangle(X1,Y1,Z1,X2,Y2,Z2) 的点是角点而非半边长 → 用 CreateCornerRectangle。
- 体积单位 m^3。基准面中文名"前视基准面"，草图名"草图N"。
"""
import os
import sys

TMP = os.path.dirname(os.path.abspath(__file__))
if TMP not in sys.path:
    sys.path.insert(0, TMP)

import pythoncom
pythoncom.CoInitialize()

import win32com.client as wc
import swnsis


def connect(visible: bool = True):
    """附着或冷启动 SolidWorks，返回 (sw, IModelDoc2 未建文档前的 ISldWorks 类型化对象)。"""
    sw = wc.Dispatch("SldWorks.Application")
    if visible:
        sw.Visible = True
    styp = swnsis.ISldWorks(
        sw._oleobj_.QueryInterface(swnsis.ISldWorks.CLSID, pythoncom.IID_IDispatch))
    return sw, styp


def new_part(sw, styp):
    """关闭全部旧文档并新建空白零件，返回五件套类型化接口。"""
    for _ in range(10):
        if int(sw.GetDocumentCount) == 0:
            break
        d = styp.GetFirstDocument()
        try:
            sw.CloseDoc(str(d.GetTitle))
        except Exception:
            pass
    model = styp.INewPart()
    if model is None:
        raise RuntimeError("INewPart failed")
    mtyp = swnsis.IModelDoc2(
        model._oleobj_.QueryInterface(swnsis.IModelDoc2.CLSID, pythoncom.IID_IDispatch))
    ext = swnsis.IModelDocExtension(
        mtyp.Extension._oleobj_.QueryInterface(swnsis.IModelDocExtension.CLSID, pythoncom.IID_IDispatch))
    skmgr = swnsis.ISketchManager(
        mtyp.SketchManager._oleobj_.QueryInterface(swnsis.ISketchManager.CLSID, pythoncom.IID_IDispatch))
    featmgr = swnsis.IFeatureManager(
        mtyp.FeatureManager._oleobj_.QueryInterface(swnsis.IFeatureManager.CLSID, pythoncom.IID_IDispatch))
    return model, mtyp, ext, skmgr, featmgr


def select(ext, name, typ):
    """按候选名选中（自动试中英文本地化名）。"""
    for cand in (name, {"前视基准面": "Front Plane", "草图": "Sketch"}.get(name, "")):
        if cand and ext.SelectByID2(cand, typ, 0, 0, 0, False, 0, None, 0):
            return True
    return False


def box(mtyp, ext, skmgr, featmgr, w_mm, h_mm, d_mm, plane="前视基准面"):
    """在指定基准面上建 w×h×d 毫米长方体，返回体积(m^3)。"""
    mtyp.ClearSelection2(True)
    assert ext.SelectByID2(plane, "PLANE", 0, 0, 0, False, 0, None, 0), f"plane {plane} not found"
    skmgr.InsertSketch(True)
    w, h = w_mm / 1000, h_mm / 1000
    skmgr.CreateCornerRectangle(-w / 2, -h / 2, 0.0, w / 2, h / 2, 0.0)
    skmgr.InsertSketch(True)
    mtyp.EditRebuild3()
    mtyp.ClearSelection2(True)
    # 选最新草图：从 草图9 往下找到存在的最大编号
    ok = False
    for i in range(9, 0, -1):
        for pre in ("草图", "Sketch"):
            if ext.SelectByID2(f"{pre}{i}", "SKETCH", 0, 0, 0, False, 0, None, 0):
                ok = True
                break
        if ok:
            break
    assert ok, "no sketch to extrude"
    featmgr.FeatureExtrusion3(
        True, False, False, 0, 0, d_mm / 1000, 0.0,
        False, False, False, False, 0, 0,
        False, False, False, False, True, True, True, 0, 0, False)
    mtyp.EditRebuild3()
    mp = ext.CreateMassProperty()
    return float(mp.Volume)


__all__ = ["connect", "new_part", "select", "box", "wc", "swnsis"]
