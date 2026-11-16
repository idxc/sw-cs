#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""从「已审核工作簿」重生成 Skill 的三份口径文件。

用法：
    python3 scripts/refresh.py /path/to/your-sop.xlsx

为什么要有这个脚本：手抄 276 个字段必然漂移。口径只在工作簿里改，
改完跑一次，references/ 下三份文件自动重生成。
"""
import json
import os
import re
import sys

import openpyxl

FIELDS = ["适用对象", "触发条件", "目的", "执行人", "操作步骤",
          "需要确认的信息", "推荐话术", "客户可能的回复与应对",
          "停止或转人工条件", "需要记录的信息", "注意事项与承诺边界", "依据状态"]

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
REF = os.path.join(ROOT, "references")


def norm(v):
    if v is None:
        return ""
    if isinstance(v, bool):
        return "是" if v else "否"
    if isinstance(v, (int, float)):
        return str(int(v)) if float(v).is_integer() else str(v)
    return str(v).strip().replace("\u3000", " ")


def as_date(v):
    if isinstance(v, (int, float)) and 40000 < v < 60000:
        from datetime import datetime, timedelta
        return (datetime(1899, 12, 30) + timedelta(days=float(v))).strftime("%Y-%m-%d")
    return norm(v)


def load(path):
    wb = openpyxl.load_workbook(path, data_only=True)

    ws = wb["01-待确认规则"]
    rules = []
    for r in range(4, ws.max_row + 1):
        rid = norm(ws.cell(row=r, column=1).value)
        if not rid.startswith("R-"):
            continue
        rules.append({
            "id": rid,
            "name": norm(ws.cell(row=r, column=2).value),
            "cur": norm(ws.cell(row=r, column=3).value),
            "where": norm(ws.cell(row=r, column=4).value),
            "text": norm(ws.cell(row=r, column=5).value),
            "owner": norm(ws.cell(row=r, column=6).value),
            "from": as_date(ws.cell(row=r, column=7).value),
            "sop": norm(ws.cell(row=r, column=10).value),
        })

    ws = wb["02-SOP场景卡"]
    sops, cur = [], None
    for r in range(4, ws.max_row + 1):
        a = norm(ws.cell(row=r, column=1).value)
        if not a:
            continue
        if re.match(r"^PSOP-\d+", a):
            m = re.match(r"^(PSOP-\d+)\s*(.+?)\s*｜\s*(.+?)\s*｜", a)
            cur = {"id": m.group(1) if m else a[:8],
                   "title": m.group(2) if m else a,
                   "stage": m.group(3) if m else "",
                   "fields": {}, "fix": {}}
            sops.append(cur)
            continue
        if cur is None:
            continue
        if a in FIELDS:
            cur["fields"][a] = norm(ws.cell(row=r, column=2).value)
            cur["fix"][a] = norm(ws.cell(row=r, column=3).value)

    ws = wb["04-话术库22场景"]
    scenes, cur = [], None
    cols = {"适用情况": 3, "标准话术": 4, "简短版": 5, "需进一步沟通时": 6,
            "可替换信息": 7, "不建议使用的表达": 8, "注意事项": 9}
    for r in range(4, ws.max_row + 1):
        no = norm(ws.cell(row=r, column=1).value)
        name = norm(ws.cell(row=r, column=2).value)
        if not name:
            continue
        item = {"no": no, "name": name, "rows": {}, "fix": ""}
        for k, c in cols.items():
            item["rows"][k] = norm(ws.cell(row=r, column=c).value)
        item["fix"] = norm(ws.cell(row=r, column=10).value)
        scenes.append(item)

    return rules, sops, scenes


def w(path, text):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        f.write(text)
    print("wrote", os.path.relpath(path, ROOT), len(text), "chars")


def main():
    if len(sys.argv) < 2:
        print("用法：python3 scripts/refresh.py <your-sop.xlsx 的完整路径>", file=sys.stderr)
        print("说明：本脚本不猜路径。工作簿每次更新后跑一次，references/ 下三份文件自动重生成。", file=sys.stderr)
        sys.exit(2)
    src = os.path.expanduser(sys.argv[1])
    if not os.path.isfile(src):
        print(f"找不到文件：{src}", file=sys.stderr)
        sys.exit(2)
    rules, sops, scenes = load(src)

    # ---- 01 口径 ----
    L = ["# 口径卡（20 项规则）", "",
         "> 来源：示例口径，由 `scripts/refresh.py` 从「SOP 工作簿」生成。**使用前请替换为本公司现行制度。**",
         "> **本文件由 `scripts/refresh.py` 生成，不要手改。** 改口径请改工作簿后重跑。", "",
         "## 用法", "",
         "先查本文件。查到了直接用；查到「L3 请示」的边界，才转升级。", "",
         "---", ""]
    for x in rules:
        L += [f"## {x['id']} {x['name']}", "", x["text"] or "（未填）", ""]
        if x["sop"]:
            L += [f"对应 SOP：{x['sop']}", ""]
        L += ["---", ""]
    w(os.path.join(REF, "01-口径卡.md"), "\n".join(L))

    # ---- 02 场景路由 ----
    L = ["# 场景路由（23 条 PSOP）", "",
         "> 客户说了一句话，先在这里定位属于哪条，再去 `03-话术库.md` 找对应说法。",
         "> **本文件由 `scripts/refresh.py` 生成，不要手改。**", ""]
    stages = {}
    for s in sops:
        stages.setdefault(s["stage"] or "未分组", []).append(s)
    for stage, items in stages.items():
        L += [f"## {stage}", ""]
        for s in items:
            trig = s["fix"].get("触发条件") or s["fields"].get("触发条件", "")
            L += [f"### {s['id']} {s['title']}", "", f"**触发**：{trig}", ""]
            for k in ("操作步骤", "需要确认的信息", "停止或转人工条件",
                      "注意事项与承诺边界"):
                v = s["fix"].get(k) or s["fields"].get(k, "")
                if v:
                    L += [f"**{k}**：{v}", ""]
            L += ["---", ""]
    w(os.path.join(REF, "02-场景路由.md"), "\n".join(L))

    # ---- 03 话术库 ----
    L = ["# 话术库（22 场景）", "",
         "> 修订版优先于原版。改语气请改工作簿第 4 表，重跑 `scripts/refresh.py`。",
         "> **本文件由脚本生成，不要手改。**", ""]
    for sc in scenes:
        L += [f"## {sc['no']}. {sc['name']}", ""]
        for k in ("适用情况", "标准话术", "简短版", "需进一步沟通时",
                  "可替换信息", "不建议使用的表达", "注意事项"):
            v = sc["rows"].get(k, "")
            if v:
                L += [f"**{k}**", "", v, ""]
        if sc["fix"]:
            L += ["> **负责人修订版（优先使用）**", "", sc["fix"], ""]
        L += ["---", ""]
    w(os.path.join(REF, "03-话术库.md"), "\n".join(L))

    print(f"规则 {len(rules)} / SOP {len(sops)} / 话术 {len(scenes)}")


if __name__ == "__main__":
    main()
