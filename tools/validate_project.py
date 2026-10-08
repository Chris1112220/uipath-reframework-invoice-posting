"""Static checks for the UiPath project, runnable without Studio (e.g. in CI).

    pip install lxml
    python tools/validate_project.py

Checks:
  * every .xaml is well-formed XML with an <Activity> root and x:Class matching the file name
  * sap2010:WorkflowViewState.IdRef values are unique within each file
  * every Invoke Workflow File target exists, and every argument passed matches an
    argument declared by the target (name, direction and type)
  * project.json / entry-points.json are valid JSON and every file they reference exists
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

from lxml import etree

ROOT = Path(__file__).resolve().parents[1]
NS = {
    "x": "http://schemas.microsoft.com/winfx/2006/xaml",
    "ui": "http://schemas.uipath.com/workflow/activities",
    "sap2010": "http://schemas.microsoft.com/netfx/2010/xaml/activities/presentation",
    "a": "http://schemas.microsoft.com/netfx/2009/xaml/activities",
}
IDREF = "{%s}WorkflowViewState.IdRef" % NS["sap2010"]
XKEY = "{%s}Key" % NS["x"]
XCLASS = "{%s}Class" % NS["x"]

errors: list[str] = []


def err(msg: str) -> None:
    errors.append(msg)


def members(tree) -> dict[str, tuple[str, str]]:
    """name -> (direction, type) from x:Members."""
    out = {}
    for prop in tree.iterfind(".//x:Members/x:Property", NS):
        typ = prop.get("Type")  # e.g. InArgument(x:String)
        direction, inner = typ.split("(", 1)
        out[prop.get("Name")] = (direction.replace("Argument", ""), inner[:-1])
    return out


def main() -> int:
    xamls = sorted(p for p in ROOT.rglob("*.xaml") if ".local" not in p.parts)
    parsed = {}
    for path in xamls:
        rel = path.relative_to(ROOT).as_posix()
        try:
            tree = etree.parse(str(path))
        except etree.XMLSyntaxError as exc:
            err(f"{rel}: not well-formed XML: {exc}")
            continue
        root = tree.getroot()
        if etree.QName(root).localname != "Activity":
            err(f"{rel}: root element is not <Activity>")
        if root.get(XCLASS) != path.stem:
            err(f"{rel}: x:Class '{root.get(XCLASS)}' does not match file name '{path.stem}'")
        ids = [e.get(IDREF) for e in root.iter() if isinstance(e.tag, str) and e.get(IDREF)]
        ids += [e.text for e in root.iter("{%s}WorkflowViewState.IdRef" % NS["sap2010"])]
        dupes = {i for i in ids if ids.count(i) > 1}
        if dupes:
            err(f"{rel}: duplicate IdRef values {sorted(dupes)}")
        parsed[rel] = tree

    invoke_count = 0
    for rel, tree in parsed.items():
        for inv in tree.iterfind(".//ui:InvokeWorkflowFile", NS):
            invoke_count += 1
            target = inv.get("WorkflowFileName").replace("\\", "/")
            if target not in parsed:
                err(f"{rel}: invokes missing workflow '{target}'")
                continue
            declared = members(parsed[target])
            passed = {}
            for arg in inv.iterfind("ui:InvokeWorkflowFile.Arguments/*", NS):
                direction = etree.QName(arg).localname.replace("Argument", "")
                passed[arg.get(XKEY)] = (direction, arg.get("{%s}TypeArguments" % NS["x"]))
            for name, sig in passed.items():
                if name not in declared:
                    err(f"{rel}: passes unknown argument '{name}' to {target}")
                elif declared[name] != sig:
                    err(f"{rel}: argument '{name}' to {target} is {sig}, target declares {declared[name]}")
            for name, (direction, _) in declared.items():
                if direction == "In" and name not in passed:
                    err(f"{rel}: does not pass In argument '{name}' to {target}")

    for name in ("project.json", "entry-points.json", "project.uiproj"):
        try:
            json.loads((ROOT / name).read_text(encoding="utf-8"))
        except Exception as exc:  # noqa: BLE001
            err(f"{name}: invalid JSON: {exc}")
    project = json.loads((ROOT / "project.json").read_text(encoding="utf-8"))
    referenced = [project["main"], project["runtimeOptions"].get("exceptionHandlerWorkflow", "")]
    referenced += [e["filePath"] for e in project.get("entryPoints", [])]
    referenced += [f["fileName"] for f in project["designOptions"].get("fileInfoCollection", [])]
    entry = json.loads((ROOT / "entry-points.json").read_text(encoding="utf-8"))
    referenced += [e["filePath"] for e in entry["entryPoints"]]
    for ref in filter(None, referenced):
        if not (ROOT / ref.replace("\\", "/")).is_file():
            err(f"project files reference missing file '{ref}'")

    print(f"Checked {len(parsed)} workflows, {invoke_count} Invoke Workflow File calls, "
          f"{len(referenced)} project.json/entry-points references.")
    for e in errors:
        print("ERROR:", e)
    print("OK" if not errors else f"{len(errors)} problem(s) found")
    return 1 if errors else 0


if __name__ == "__main__":
    sys.exit(main())
