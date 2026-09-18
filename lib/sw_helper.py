r"""SolidWorks automation helper (Windows COM; self-contained release copy).

Verified on SolidWorks 2024 SP5 via win32com dynamic dispatch + makepy wrapper
(swnsis.py). Key semantics: RevisionNumber/GetTitle/GetType/Volume are PROPERTIES;
SelectByID2 needs the typed wrapper; blank docs come from INewPart(); extrude uses
CreateCornerRectangle (not Center); mass-property Volume is m^3.
"""
import os
import sys

TMP = os.path.dirname(os.path.abspath(__file__))
if TMP not in sys.path:
    sys.path.insert(0, TMP)

import pythoncom
pythoncom.CoInitialize()

import win32com.client as wc
import swnsis          # SolidWorks 2024 typelib wrapper (makepy)


def connect(visible: bool = True):
    """(ascii repair.)"""
    sw = wc.Dispatch("SldWorks.Application")
    if visible:
        sw.Visible = True
    styp = swnsis.ISldWorks(
        sw._oleobj_.QueryInterface(swnsis.ISldWorks.CLSID, pythoncom.IID_IDispatch))
    return sw, styp


def new_part(sw, styp):
    """(ascii repair.)"""
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
    """Select by name; Chinese localized plane/sketch names map to English aliases."""
    aliases = {
        "\u524d\u89c6\u57fa\u51c6\u9762": ["Front Plane"],
        "\u8349\u56fe": ["Sketch"],
    }
    cands = [name] + aliases.get(name, [])
    for cand in cands:
        if ext.SelectByID2(cand, typ, 0, 0, 0, False, 0, None, 0):
            return True
    return False


def box(mtyp, ext, skmgr, featmgr, w_mm, h_mm, d_mm, plane="\u524d\u89c6\u57fa\u51c6\u9762"):
    """(ascii repair.)"""
    mtyp.ClearSelection2(True)
    assert ext.SelectByID2(plane, "PLANE", 0, 0, 0, False, 0, None, 0), f"plane {plane} not found"
    skmgr.InsertSketch(True)
    w, h = w_mm / 1000, h_mm / 1000
    skmgr.CreateCornerRectangle(-w / 2, -h / 2, 0.0, w / 2, h / 2, 0.0)
    skmgr.InsertSketch(True)
    mtyp.EditRebuild3()
    mtyp.ClearSelection2(True)
    ok = False
    for i in range(9, 0, -1):
        for pre in ("\u8349\u56fe", "Sketch"):
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
