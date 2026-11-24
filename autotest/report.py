"""测试报告生成：HTML + Excel。"""

from __future__ import annotations

import base64
import html
from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .models import Project, status_label


STATUS_COLORS = {
    "pass": ("DCFCE7", "166534", "通过"),
    "fail": ("FEE2E2", "991B1B", "不通过"),
    "not_configured": ("FEF3C7", "92400E", "未配置"),
    "skipped": ("E2E8F0", "334155", "跳过"),
    "error": ("FCE7F3", "9D174D", "执行错误"),
    "pending": ("E0F2FE", "075985", "待运行"),
}


def _esc(value: Any) -> str:
    return html.escape(str(value or ""), quote=True)


def _screenshot_data_uri(run_dir: Path, relative_path: str) -> str:
    if not relative_path:
        return ""
    image = run_dir / relative_path
    if not image.exists():
        return ""
    try:
        data = base64.b64encode(image.read_bytes()).decode("ascii")
        return f"data:image/png;base64,{data}"
    except OSError:
        return ""


def _status_badge(status: str) -> str:
    return (
        f'<span class="badge badge-{_esc(status)}">'
        f"{_esc(status_label(status))}</span>"
    )


def generate_reports(
    project: Project, metadata: dict[str, Any], run_dir: Path
) -> tuple[Path, Path]:
    run_dir.mkdir(parents=True, exist_ok=True)
    html_path = run_dir / "report.html"
    xlsx_path = run_dir / "report.xlsx"
    _write_html(project, metadata, run_dir, html_path)
    _write_xlsx(project, metadata, xlsx_path)
    return html_path, xlsx_path


def _write_html(
    project: Project,
    metadata: dict[str, Any],
    run_dir: Path,
    path: Path,
) -> None:
    counts = dict(metadata.get("counts", {}))
    total = int(metadata.get("total", project.case_count) or 0)
    pass_count = int(counts.get("pass", 0))
    fail_count = int(counts.get("fail", 0))
    unconfigured = int(counts.get("not_configured", 0))
    skipped = int(counts.get("skipped", 0)) + int(counts.get("error", 0))

    failed_cases = [
        (req, case)
        for req in project.requirements
        for case in req.cases
        if case.status in ("fail", "error")
    ]

    failure_cards = ""
    if failed_cases:
        items = []
        for req, case in failed_cases:
            shot = _screenshot_data_uri(run_dir, case.screenshot)
            image_html = (
                f'<div class="evidence"><img src="{shot}" alt="失败截图"></div>'
                if shot
                else ""
            )
            items.append(
                f"""
                <article class="failure-card">
                  <div class="failure-head">
                    <strong>{_esc(req.seq)} · {_esc(req.module)} / {_esc(req.point)}</strong>
                    <span class="badge badge-fail">不通过</span>
                  </div>
                  <h3>{_esc(case.name)}</h3>
                  <dl>
                    <dt>需求/预期</dt><dd>{_esc(case.expected)}</dd>
                    {f'<dt>需规备注</dt><dd>{_esc(req.remark)}</dd>' if req.remark else ''}
                    <dt>实际表现</dt><dd>{_esc(case.actual or case.error)}</dd>
                  </dl>
                  {image_html}
                </article>
                """
            )
        failure_cards = """
        <section class="section">
          <h2>不满足需求的项目</h2>
          <p class="muted">以下用例的实际表现与预期不一致，需要优先修复或进一步人工复核。</p>
          <div class="failure-grid">""" + "".join(items) + "</div></section>"
    else:
        if pass_count == 0 and unconfigured > 0:
            status_message = f"""
            <div class="warn-state">
              本次没有执行任何可自动判定的用例。{unconfigured} 条用例被标记为“未配置”，
              因此不能据此判断需求满足或不满足。
            </div>"""
        elif unconfigured > 0:
            status_message = f"""
            <div class="warn-state">
              已执行用例中没有发现不通过项，但仍有 {unconfigured} 条用例未配置步骤，
              存在未验证的需求范围。
            </div>"""
        else:
            status_message = '<div class="ok-state">本次已执行用例中没有发现不通过项。</div>'
        failure_cards = f"""
        <section class="section">
          <h2>不满足需求的项目</h2>
          {status_message}
        </section>"""

    rows = []
    for req in project.requirements:
        for case in req.cases:
            shot = _screenshot_data_uri(run_dir, case.screenshot)
            evidence = (
                f'<details><summary>查看截图</summary><img class="thumb" src="{shot}" alt="截图"></details>'
                if shot
                else "—"
            )
            error = f'<div class="error-text">{_esc(case.error)}</div>' if case.error else ""
            rows.append(
                f"""
                <tr data-status="{_esc(case.status)}">
                  <td>{_esc(req.seq)}</td>
                  <td><div>{_esc(req.module)}</div><small>{_esc(req.function)}</small></td>
                  <td>{_esc(req.point)}</td>
                  <td>{_esc(req.remark)}</td>
                  <td><div>{_esc(case.name)}</div><small>{_esc(case.description)}</small></td>
                  <td>{_status_badge(case.status)}</td>
                  <td>{_esc(case.expected)}</td>
                  <td>{_esc(case.actual)}</td>
                  <td>{evidence}{error}</td>
                  <td class="nowrap">{case.duration_s:.2f}s</td>
                </tr>
                """
            )

    generated = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    document = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>测试报告 - {_esc(project.name)}</title>
<style>
:root {{
  --bg: #f4f7fb; --surface: #fff; --text: #1f2937; --muted: #667085;
  --line: #dbe3ee; --primary: #2563eb; --danger: #b42318;
}}
* {{ box-sizing: border-box; }}
body {{ margin: 0; background: var(--bg); color: var(--text);
  font-family: "Microsoft YaHei UI", "Microsoft YaHei", system-ui, sans-serif;
  line-height: 1.6; }}
.page {{ max-width: 1500px; margin: 0 auto; padding: 32px 24px 64px; }}
.hero {{ background: linear-gradient(135deg, #0f172a, #1e3a8a);
  color: white; border-radius: 18px; padding: 30px 34px; box-shadow: 0 18px 40px rgba(15,23,42,.18); }}
.hero h1 {{ margin: 0 0 6px; font-size: 30px; letter-spacing: -.02em; }}
.hero p {{ margin: 2px 0; color: #cbd5e1; font-size: 13px; }}
.meta {{ display: flex; flex-wrap: wrap; gap: 8px 24px; margin-top: 18px; font-size: 13px; color: #dbeafe; }}
.summary {{ display: grid; grid-template-columns: repeat(5, minmax(120px, 1fr)); gap: 12px; margin: 18px 0 24px; }}
.card {{ background: var(--surface); border: 1px solid var(--line); border-radius: 14px; padding: 18px 20px; }}
.card .num {{ font-size: 30px; font-weight: 800; line-height: 1.15; }}
.card .label {{ color: var(--muted); font-size: 12px; font-weight: 600; }}
.pass .num {{ color: #047857; }} .fail .num {{ color: var(--danger); }}
.unconfigured .num {{ color: #b45309; }} .total .num {{ color: #1d4ed8; }}
.section {{ background: var(--surface); border: 1px solid var(--line); border-radius: 14px;
  padding: 22px; margin-top: 18px; }}
.section h2 {{ margin: 0 0 4px; font-size: 19px; }}
.muted {{ color: var(--muted); }}
.failure-grid {{ display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 14px; margin-top: 14px; }}
.failure-card {{ border: 1px solid #fecaca; background: #fffafa; border-radius: 12px; padding: 16px; }}
.failure-head {{ display: flex; justify-content: space-between; gap: 10px; align-items: flex-start; }}
.failure-card h3 {{ margin: 6px 0 8px; font-size: 15px; }}
.failure-card dl {{ margin: 0; font-size: 13px; }}
.failure-card dt {{ color: var(--muted); font-weight: 700; margin-top: 7px; }}
.failure-card dd {{ margin: 0; overflow-wrap: anywhere; }}
.evidence img {{ max-width: 100%; border: 1px solid var(--line); border-radius: 8px; margin-top: 10px; }}
.ok-state {{ margin-top: 12px; background: #ecfdf5; color: #065f46; border: 1px solid #a7f3d0;
  border-radius: 10px; padding: 14px 16px; }}
.warn-state {{ margin-top: 12px; background: #fffbeb; color: #92400e; border: 1px solid #fde68a;
  border-radius: 10px; padding: 14px 16px; }}
.badge {{ display: inline-block; border-radius: 999px; padding: 3px 10px; font-size: 12px; font-weight: 700; white-space: nowrap; }}
.badge-pass {{ background: #dcfce7; color: #166534; }}
.badge-fail, .badge-error {{ background: #fee2e2; color: #991b1b; }}
.badge-not_configured {{ background: #fef3c7; color: #92400e; }}
.badge-skipped {{ background: #e2e8f0; color: #334155; }}
.badge-pending {{ background: #e0f2fe; color: #075985; }}
.toolbar {{ display: flex; flex-wrap: wrap; gap: 8px; align-items: center; margin: 18px 0 10px; }}
.toolbar button {{ border: 1px solid var(--line); background: #fff; border-radius: 8px; padding: 7px 13px;
  cursor: pointer; font: inherit; font-size: 13px; }}
.toolbar button:hover {{ border-color: #93c5fd; color: var(--primary); }}
.toolbar .spacer {{ flex: 1; }}
.table-wrap {{ overflow: auto; border: 1px solid var(--line); border-radius: 12px; background: #fff; }}
table {{ border-collapse: collapse; width: 100%; min-width: 1320px; font-size: 12.5px; }}
th {{ position: sticky; top: 0; z-index: 2; background: #f8fafc; text-align: left; color: #475569;
  padding: 11px 10px; border-bottom: 1px solid var(--line); white-space: nowrap; }}
td {{ vertical-align: top; padding: 11px 10px; border-bottom: 1px solid #edf2f7; overflow-wrap: anywhere; }}
td small {{ color: var(--muted); display: block; font-size: 11px; margin-top: 2px; }}
tr:hover td {{ background: #fbfdff; }}
.nowrap {{ white-space: nowrap; }}
.error-text {{ color: #991b1b; font-size: 11px; margin-top: 5px; }}
.thumb {{ display: block; max-width: 320px; max-height: 220px; margin-top: 8px; border: 1px solid var(--line); border-radius: 6px; }}
.footer {{ color: var(--muted); font-size: 12px; text-align: center; margin-top: 24px; }}
@media (max-width: 900px) {{
  .summary {{ grid-template-columns: repeat(2, 1fr); }}
  .failure-grid {{ grid-template-columns: 1fr; }}
}}
@media print {{
  body {{ background: #fff; }}
  .page {{ max-width: none; padding: 0; }}
  .toolbar {{ display: none; }}
  .table-wrap {{ overflow: visible; }}
  th {{ position: static; }}
  .failure-card {{ break-inside: avoid; }}
}}
</style>
</head>
<body>
<div class="page">
  <header class="hero">
    <h1>测试报告：{_esc(project.name)}</h1>
    <p>导入文件：{_esc(project.source_excel or metadata.get("source_excel", ""))}</p>
    <p>数据表：{_esc(project.source_sheet or "-")}　生成时间：{_esc(generated)}</p>
    <div class="meta">
      <span>待测网址：{_esc(metadata.get("base_url", ""))}</span>
      <span>浏览器：{_esc(metadata.get("browser", ""))}</span>
      <span>开始：{_esc(metadata.get("started_at", ""))}</span>
      <span>结束：{_esc(metadata.get("finished_at", ""))}</span>
    </div>
  </header>

  <section class="summary">
    <div class="card total"><div class="num">{total}</div><div class="label">用例总数</div></div>
    <div class="card pass"><div class="num">{pass_count}</div><div class="label">通过</div></div>
    <div class="card fail"><div class="num">{fail_count}</div><div class="label">不通过</div></div>
    <div class="card unconfigured"><div class="num">{unconfigured}</div><div class="label">未配置步骤</div></div>
    <div class="card"><div class="num">{skipped}</div><div class="label">跳过 / 错误</div></div>
  </section>

  {failure_cards}

  <section class="section">
    <h2>全部用例明细</h2>
    <p class="muted">“未配置”表示该条需求已导入，但还没有可执行的页面步骤；它不能被算作通过。</p>
    <div class="toolbar">
      <button onclick="filterRows('all', this)">全部</button>
      <button onclick="filterRows('pass', this)">通过</button>
      <button onclick="filterRows('fail', this)">不通过</button>
      <button onclick="filterRows('not_configured', this)">未配置</button>
      <button onclick="filterRows('skipped', this)">跳过</button>
      <span class="spacer"></span>
      <button onclick="window.print()">打印 / 另存 PDF</button>
    </div>
    <div class="table-wrap">
      <table id="detailTable">
        <thead><tr>
          <th>序号</th><th>模块 / 功能</th><th>功能点</th><th>需规备注</th><th>用例</th><th>状态</th>
          <th>预期结果</th><th>实际表现</th><th>证据 / 错误</th><th>耗时</th>
        </tr></thead>
        <tbody>{''.join(rows)}</tbody>
      </table>
    </div>
  </section>

  <footer class="footer">由桌面自动化测试校验系统生成 · {_esc(generated)}</footer>
</div>
<script>
function filterRows(status, button) {{
  document.querySelectorAll('#detailTable tbody tr').forEach(function (row) {{
    row.style.display = (status === 'all' || row.dataset.status === status) ? '' : 'none';
  }});
  document.querySelectorAll('.toolbar button').forEach(function (btn) {{ btn.classList.remove('active'); }});
  if (button) {{ button.classList.add('active'); }}
}}
</script>
</body>
</html>"""
    path.write_text(document, encoding="utf-8")


def _write_xlsx(
    project: Project, metadata: dict[str, Any], path: Path
) -> None:
    wb = Workbook()
    summary = wb.active
    summary.title = "汇总"
    counts = dict(metadata.get("counts", {}))
    summary.append(["项目", project.name])
    summary.append(["导入文件", project.source_excel])
    summary.append(["数据表", project.source_sheet])
    summary.append(["待测网址", metadata.get("base_url", "")])
    summary.append(["浏览器", metadata.get("browser", "")])
    summary.append(["开始时间", metadata.get("started_at", "")])
    summary.append(["结束时间", metadata.get("finished_at", "")])
    summary.append([])
    summary.append(["指标", "数量"])
    summary.append(["用例总数", metadata.get("total", project.case_count)])
    for key in ("pass", "fail", "not_configured", "skipped", "error", "pending"):
        summary.append([status_label(key), counts.get(key, 0)])
    summary.column_dimensions["A"].width = 18
    summary.column_dimensions["B"].width = 80
    summary["A1"].font = Font(bold=True, size=14, color="1E3A8A")

    detail = wb.create_sheet("明细")
    headers = [
        "序号", "系统", "模块", "功能", "功能点", "需规备注", "用例名称", "状态",
        "预期结果", "实际表现", "错误信息", "截图", "耗时(秒)", "来源表", "来源行",
    ]
    detail.append(headers)
    thin = Side(style="thin", color="CBD5E1")
    header_fill = PatternFill("solid", fgColor="1E3A8A")
    for cell in detail[1]:
        cell.fill = header_fill
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(bottom=thin)

    for req in project.requirements:
        for case in req.cases:
            detail.append(
                [
                    req.seq,
                    req.system,
                    req.module,
                    req.function,
                    req.point,
                    req.remark,
                    case.name,
                    status_label(case.status),
                    case.expected,
                    case.actual,
                    case.error,
                    case.screenshot,
                    case.duration_s,
                    req.source_sheet,
                    req.source_row,
                ]
            )

    widths = [8, 18, 16, 18, 20, 26, 28, 10, 45, 55, 45, 30, 10, 14, 10]
    for index, width in enumerate(widths, start=1):
        detail.column_dimensions[get_column_letter(index)].width = width
    detail.freeze_panes = "A2"
    detail.auto_filter.ref = f"A1:O{detail.max_row}"

    for row in detail.iter_rows(min_row=2, max_row=detail.max_row):
        status = str(row[7].value or "")
        for key, (fill_color, font_color, label) in STATUS_COLORS.items():
            if status == label:
                row[7].fill = PatternFill("solid", fgColor=fill_color)
                row[7].font = Font(color=font_color, bold=True)
                break
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = Border(bottom=Side(style="hair", color="E2E8F0"))

    wb.save(path)
