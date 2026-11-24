"""构建并运行“志愿星图”15 条验收用例。"""

from pathlib import Path
import sys

sys.path.insert(0, r"D:\AutoTestWorkbench")

from autotest.excel_import import export_project_plan, import_workbook
from autotest.models import Project, Requirement, Step, TestCase
from autotest.runner import RunOptions, WebTestRunner
from autotest.storage import AppStorage

APP = Path(r"D:\高考志愿填报助手\index.html")
URI = APP.as_uri()
PLAN = Path(r"D:\AutoTestWorkbench\inputs\documents\志愿星图_自动化验收测试计划.xlsx")
SYSTEM = "志愿星图"
MODULE = "验收测试"


def login_steps():
    return [
        Step(action="goto", target=URI, description="打开离线应用"),
        Step(action="fill", target="#login-id", value="{{username}}", description="输入测试账号"),
        Step(action="fill", target="#login-password", value="{{password}}", description="输入密码"),
        Step(action="click", target=".login-submit", description="点击进入系统"),
        Step(action="assert_visible", target=".app-shell", description="首页外壳出现"),
        Step(action="assert_url", value="#/home", description="进入首页路由"),
    ]


def recommend_steps(score="600"):
    return login_steps() + [
        Step(action="click", target='a[data-nav="recommend"]', description="进入智能填报"),
        Step(action="wait", target="#rec-score", description="等待填报表单"),
        Step(action="fill", target="#rec-score", value=score, description="输入分数"),
        Step(action="click", target='form.query-form button[type="submit"]', description="生成专业志愿"),
    ]


def requirement(req_id, point, actual, row, case_name, expected, steps, module=MODULE, function="验收"):
    return Requirement(
        req_id=req_id,
        seq=req_id,
        system=SYSTEM,
        module=module,
        function=function,
        point=point,
        actual_features=actual,
        source_sheet="08 验收标准",
        source_row=row,
        cases=[
            TestCase(
                case_id=f"{req_id}-C01",
                requirement_id=req_id,
                name=case_name,
                expected=expected,
                steps=steps,
            )
        ],
    )


project = Project.create("志愿星图 · 高考志愿智能填报 V2.0 验收测试", str(APP), "08 验收标准")
reqs = []

reqs.append(requirement(
    "TC-001", "正确账号登录",
    "账号 250101 / 密码 0101，登录后进入首页并显示管理员1",
    6, "正确账号登录", "进入首页，顶部显示管理员1",
    login_steps() + [Step(action="assert_text", target=".user-card", value="管理员1")],
    function="登录",
))

reqs.append(requirement(
    "TC-002", "错误密码",
    "输入 250101 / 9999，提示账号或密码不正确，停留登录页",
    7, "错误密码提示", "提示“账号或密码不正确”，停留登录页",
    [
        Step(action="goto", target=URI),
        Step(action="fill", target="#login-id", value="250101"),
        Step(action="fill", target="#login-password", value="9999"),
        Step(action="click", target=".login-submit"),
        Step(action="assert_text", target=".login-error", value="账号或密码不正确"),
        Step(action="assert_visible", target="#login-id"),
    ],
    function="登录",
))

reqs.append(requirement(
    "TC-003", "格式校验",
    "5 位账号和 3 位密码分别提示位数错误",
    8, "账号和密码格式校验", "分别提示账号/密码位数错误",
    [
        Step(action="goto", target=URI),
        Step(action="fill", target="#login-id", value="25010"),
        Step(action="fill", target="#login-password", value="0101"),
        Step(action="click", target=".login-submit"),
        Step(action="assert_text", target=".login-error", value="账号应为 6 位数字"),
        Step(action="fill", target="#login-id", value="250101"),
        Step(action="fill", target="#login-password", value="010"),
        Step(action="click", target=".login-submit"),
        Step(action="assert_text", target=".login-error", value="密码应为 4 位数字"),
    ],
    function="登录",
))

reqs.append(requirement(
    "TC-004", "生成 96 条",
    "河南、物理类、600 分，生成 96 条，每条含院校/专业组/专业",
    9, "生成 96 个专业志愿", "结果列表 96 条，每条含院校/专业组/专业",
    recommend_steps() + [
        Step(action="assert_count", target=".rec-row", value="96"),
        Step(action="assert_count", target=".rec-school-name", value="96"),
        Step(action="assert_count", target=".rec-major-name", value="96"),
        Step(action="assert_count", target=".rec-major-group", value="96"),
    ],
    function="智能填报",
))

reqs.append(requirement(
    "TC-005", "概率排序",
    "概率非递减，首条约 5%，末条约 95%",
    10, "录取概率梯度", "概率非递减，首条 5%，末条 95%",
    recommend_steps() + [
        Step(action="assert_text", target=".rec-row:first-child .prob-value", value="5%"),
        Step(action="assert_text", target=".rec-row:last-child .prob-value", value="95%"),
        Step(action="assert_sorted", target=".rec-row .prob-value", value="asc"),
    ],
    function="智能填报",
))

reqs.append(requirement(
    "TC-006", "每校限 2 个",
    "常规情况下同一院校出现不超过 2 次",
    11, "同一院校最多 2 个专业", "同一院校出现不超过 2 次",
    recommend_steps() + [Step(action="assert_duplicate_max", target=".rec-school-name", value="2")],
    function="智能填报",
))

reqs.append(requirement(
    "TC-007", "筛选与搜索",
    "切换冲刺后搜索计算机，列表只保留对应结果，数量同步",
    12, "筛选冲刺并搜索计算机", "列表只保留冲刺且包含计算机的结果",
    recommend_steps() + [
        Step(action="click", target='.chip[data-tier="reach"]'),
        Step(action="fill", target="#rec-search", value="计算机"),
        Step(action="assert_count", target=".rec-row", value="4"),
        Step(action="assert_text", target=".rec-row", value="计算机"),
        Step(action="assert_text", target=".rec-row", value="冲刺"),
    ],
    function="智能填报",
))

detail_steps = login_steps() + [
    Step(action="click", target='a[data-nav="schools"]'),
    Step(action="wait", target="#lib-search"),
    Step(action="fill", target="#lib-search", value="上海大学"),
    Step(action="click", target=".school-card"),
    Step(action="assert_visible", target=".modal-backdrop"),
]

reqs.append(requirement(
    "TC-008", "分专业表",
    "院校库搜索上海大学，打开详情，显示 2025 分专业表",
    13, "上海大学分专业表", "显示专业组/专业/分数/位次，共 8 个专业",
    detail_steps + [
        Step(action="assert_text", target=".modal-backdrop", value="上海大学"),
        Step(action="assert_text", target=".modal-backdrop", value="分专业录取"),
        Step(action="assert_count", target=".major-table tbody tr", value="8"),
    ],
    function="院校详情",
))

reqs.append(requirement(
    "TC-009", "切换查询口径",
    "详情中把省份从河南切到安徽，专业表、统计、历史表刷新",
    14, "切换省份刷新详情", "专业表按安徽物理类刷新",
    detail_steps + [
        Step(action="select", target='select[data-change="detail-province"]', value="anhui"),
        Step(action="wait", value="500"),
        Step(action="assert_text", target=".modal-backdrop", value="安徽 · 物理类"),
        Step(action="assert_count", target=".major-table tbody tr", value="8"),
    ],
    function="院校详情",
))

reqs.append(requirement(
    "TC-010", "收藏持久化",
    "收藏一所院校，刷新后收藏状态保留，收藏页可见",
    15, "收藏上海大学并刷新", "刷新后收藏保留，收藏页出现上海大学",
    login_steps() + [
        Step(action="click", target='a[data-nav="schools"]'),
        Step(action="wait", target="#lib-search"),
        Step(action="fill", target="#lib-search", value="上海大学"),
        Step(action="click", target=".school-card .card-fav"),
        Step(action="assert_attribute", target=".school-card .card-fav", value="aria-pressed=true"),
        Step(action="reload"),
        Step(action="assert_visible", target=".app-shell"),
        Step(action="click", target='a[data-nav="favorites"]'),
        Step(action="assert_text", target="body", value="上海大学"),
    ],
    function="收藏",
))

reqs.append(requirement(
    "TC-011", "缺失值显示",
    "缺失数据统一显示“—”；以安徽上海大学录取人数缺失为例",
    16, "缺失值显示为破折号", "缺失字段显示“—”，不显示 0 或 null",
    detail_steps + [
        Step(action="select", target='select[data-change="detail-province"]', value="anhui"),
        Step(action="wait", value="500"),
        Step(action="assert_text", target=".modal-backdrop", value="—"),
    ],
    function="数据",
))

reqs.append(requirement(
    "TC-012", "打印/保存 PDF",
    "打印时隐藏侧栏、筛选栏和操作按钮，按 A4 横向输出专业志愿表",
    17, "打印媒体样式校验", "打印媒体下侧栏/筛选/操作按钮隐藏，摘要和列表保留",
    recommend_steps() + [
        Step(action="emulate_media", value="print"),
        Step(action="assert_hidden", target=".sidebar"),
        Step(action="assert_hidden", target=".rec-toolbar"),
        Step(action="assert_hidden", target=".rec-actions"),
        Step(action="assert_visible", target=".result-summary"),
        Step(action="assert_visible", target=".rec-row"),
    ],
    function="智能填报",
))

reqs.append(requirement(
    "TC-013", "快捷启动",
    "双击桌面快捷方式，打开无地址栏应用窗口，标题正确",
    18, "桌面快捷方式启动", "打开无地址栏应用窗口，标题为志愿星图·高考志愿智能填报",
    [],
    function="运行",
))

reqs.append(requirement(
    "TC-014", "数据统计",
    "打开首页查看统计：31 省、1286 院校、66,990 专业记录、51,931 院校记录",
    19, "首页统计与登录信息", "统计值和账号信息正确",
    login_steps() + [
        Step(action="assert_text", target=".stat-card", value="1286"),
        Step(action="assert_text", target="body", value="31"),
        Step(action="assert_text", target="body", value="66,990"),
        Step(action="assert_text", target="body", value="51,931"),
        Step(action="assert_text", target=".user-card", value="管理员1"),
    ],
    function="首页",
))

reqs.append(requirement(
    "TC-015", "不足 96 条",
    "选择西藏文科生成，按实际条数展示并提示未虚构填充",
    20, "西藏文科不足 96 条提示", "显示实际条数并提示未虚构填充",
    login_steps() + [
        Step(action="click", target='a[data-nav="recommend"]'),
        Step(action="wait", target="#rec-province"),
        Step(action="select", target="#rec-province", value="xizang"),
        Step(action="wait", value="400"),
        Step(action="select", target="#rec-track", value="arts"),
        Step(action="fill", target="#rec-score", value="400"),
        Step(action="click", target='form.query-form button[type="submit"]'),
        Step(action="assert_count", target=".rec-row", value="79"),
        Step(action="assert_text", target=".notice-bar", value="无法补足 96"),
    ],
    function="智能填报",
))

project.requirements = reqs
# 从需规里复制账号/权限，导出计划时一并写入“测试账号”sheet。
srs_project = import_workbook(r"C:\Users\Kevin\Desktop\志愿星图_高考志愿智能填报系统_需求规格说明书_V2.0.xlsx")
project.accounts = srs_project.accounts
export_project_plan(project, PLAN)
print("PLAN", PLAN)
print("CASES", project.case_count, "CONFIGURED", project.configured_case_count)

imported = import_workbook(PLAN)
print("IMPORTED", len(imported.requirements), imported.case_count, imported.configured_case_count)
storage = AppStorage()
result = WebTestRunner(storage, RunOptions(base_url=URI, headless=True, username="250101", password="0101")).run(imported)
print("accounts in imported plan:", [(a.username, a.password, a.display_name) for a in imported.accounts])
print("COUNTS", result.counts)
for req in imported.requirements:
    for case in req.cases:
        print(req.req_id, case.status, case.error[:200])
print("REPORT", result.report_html)
print("XLSX", result.report_xlsx)
