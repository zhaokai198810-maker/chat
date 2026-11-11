from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

FONT = "Microsoft YaHei"
NAVY = RGBColor(0x12, 0x2B, 0x4A)
RED = RGBColor(0xC0, 0x1E, 0x2E)
TEAL = RGBColor(0x0F, 0x76, 0x6E)
AMBER = RGBColor(0xC2, 0x7C, 0x0E)
GREY = RGBColor(0x55, 0x5F, 0x6D)
LIGHT = RGBColor(0xF2, 0xF5, 0xF9)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK = RGBColor(0x1F, 0x29, 0x37)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]
TOTAL = 21


def _font(run, size, bold=False, color=DARK):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = FONT
    rpr = run._r.get_or_add_rPr()
    for tag in ("a:ea", "a:cs"):
        el = rpr.find(qn(tag))
        if el is None:
            el = rpr.makeelement(qn(tag), {})
            rpr.append(el)
        el.set("typeface", FONT)


def text(slide, x, y, w, h, lines, size=18, color=DARK, bold=False,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, spacing=1.15):
    """lines: str or list of str / (str, dict) where dict may set size,color,bold,bullet,level."""
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(0.08)
    if isinstance(lines, str):
        lines = [lines]
    for i, item in enumerate(lines):
        if isinstance(item, tuple):
            s, opt = item
        else:
            s, opt = item, {}
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = opt.get("align", align)
        p.line_spacing = spacing
        p.space_after = Pt(opt.get("after", 6))
        level = opt.get("level", 0)
        bullet = opt.get("bullet", False)
        prefix = ""
        if bullet:
            prefix = ("•  " if level == 0 else "–  ")
            p.level = level
        r = p.add_run()
        r.text = prefix + s
        _font(r, opt.get("size", size - 2 * level), opt.get("bold", bold),
              opt.get("color", color))
    return tb


def box(slide, x, y, w, h, fill=LIGHT, line=None, shape=MSO_SHAPE.ROUNDED_RECTANGLE):
    s = slide.shapes.add_shape(shape, x, y, w, h)
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
        s.line.width = Pt(1.5)
    s.shadow.inherit = False
    if shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        s.adjustments[0] = 0.08
    return s


def card(slide, x, y, w, h, title, body, color=NAVY, size=15, title_size=17):
    box(slide, x, y, w, h, fill=WHITE, line=color)
    band = box(slide, x, y, w, Inches(0.55), fill=color, shape=MSO_SHAPE.RECTANGLE)
    text(slide, x, y, w, Inches(0.55), title, size=title_size, color=WHITE, bold=True,
         anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    text(slide, x + Inches(0.1), y + Inches(0.65), w - Inches(0.2), h - Inches(0.75),
         body, size=size)


def new_slide(title, subtitle=None, num=None):
    s = prs.slides.add_slide(BLANK)
    box(s, 0, 0, SW, Inches(1.05), fill=NAVY, shape=MSO_SHAPE.RECTANGLE)
    box(s, 0, Inches(1.05), SW, Inches(0.06), fill=RED, shape=MSO_SHAPE.RECTANGLE)
    text(s, Inches(0.5), Inches(0.12), Inches(11.5), Inches(0.8), title, size=28,
         color=WHITE, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    if subtitle:
        text(s, Inches(0.5), Inches(1.18), Inches(12.3), Inches(0.5), subtitle, size=15,
             color=GREY)
    n = len(prs.slides._sldIdLst)
    text(s, Inches(0.5), Inches(7.05), Inches(9), Inches(0.35),
         "STEMI 发病 >12 小时的再灌注决策", size=10, color=GREY)
    text(s, Inches(11.8), Inches(7.05), Inches(1.1), Inches(0.35), f"{n} / {TOTAL}",
         size=10, color=GREY, align=PP_ALIGN.RIGHT)
    return s


def source(slide, s):
    text(slide, Inches(0.5), Inches(6.62), Inches(12.3), Inches(0.42), s, size=10, color=GREY)


def table(slide, x, y, w, h, rows, col_w=None, font=12, header_fill=NAVY, first_col_bold=True):
    nr, nc = len(rows), len(rows[0])
    gt = slide.shapes.add_table(nr, nc, x, y, w, h).table
    if col_w:
        for i, cw in enumerate(col_w):
            gt.columns[i].width = Inches(cw)
    for r in range(nr):
        for c in range(nc):
            cell = gt.cell(r, c)
            val = rows[r][c]
            color = None
            if isinstance(val, tuple):
                val, color = val
            cell.text = ""
            p = cell.text_frame.paragraphs[0]
            run = p.add_run()
            run.text = val
            cell.margin_left = cell.margin_right = Inches(0.06)
            cell.margin_top = cell.margin_bottom = Inches(0.03)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            if r == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = header_fill
                _font(run, font + 1, True, WHITE)
                p.alignment = PP_ALIGN.CENTER
            else:
                cell.fill.solid()
                cell.fill.fore_color.rgb = LIGHT if r % 2 else WHITE
                _font(run, font, first_col_bold and c == 0, color or DARK)
    return gt


def pill(slide, x, y, w, h, label, fill):
    b = box(slide, x, y, w, h, fill=fill)
    text(slide, x, y, w, h, label, size=14, color=WHITE, bold=True,
         anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    return b


def arrow(slide, x, y, w, h, color=GREY):
    a = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, x, y, w, h)
    a.fill.solid()
    a.fill.fore_color.rgb = color
    a.line.fill.background()


def down_arrow(slide, x, y, w, h, color=GREY):
    a = slide.shapes.add_shape(MSO_SHAPE.DOWN_ARROW, x, y, w, h)
    a.fill.solid()
    a.fill.fore_color.rgb = color
    a.line.fill.background()


# 1 封面
s = prs.slides.add_slide(BLANK)
box(s, 0, 0, SW, SH, fill=NAVY, shape=MSO_SHAPE.RECTANGLE)
box(s, 0, Inches(4.55), SW, Inches(0.08), fill=RED, shape=MSO_SHAPE.RECTANGLE)
text(s, Inches(0.9), Inches(1.4), Inches(11.5), Inches(1.2), "STEMI 症状发作 >12 小时",
     size=48, color=WHITE, bold=True)
text(s, Inches(0.9), Inches(2.6), Inches(11.5), Inches(1.0), "该不该开？什么时候开？",
     size=40, color=RGBColor(0xFF, 0xC9, 0xC9), bold=True)
text(s, Inches(0.9), Inches(3.65), Inches(11.5), Inches(0.8),
     "晚就诊 STEMI 再灌注策略的证据与指南更新（截至 2026 年 11 月）", size=20,
     color=RGBColor(0xD6, 0xE0, 0xEC))
text(s, Inches(0.9), Inches(4.9), Inches(11.5), Inches(1.6), [
    "证据来源：2023 ESC ACS 指南 · 2025 ACC/AHA ACS 指南 · 2019 中国 STEMI 指南 · 2025 中国 PCI 指南",
    "BRAVE-2 · OAT · FAST-MI · LATE-MI（ESC 2026）· ISACS（EHJ 2026）· 2025–2026 荟萃分析",
], size=15, color=RGBColor(0xD6, 0xE0, 0xEC))
text(s, Inches(0.9), Inches(6.6), Inches(11.5), Inches(0.5),
     "仅供医学教学交流，临床决策请结合患者具体情况及所在机构规范", size=12,
     color=RGBColor(0x9F, 0xB3, 0xC8))

# 2 目录
s = new_slide("目录")
items = [
    ("01", "临床问题：谁是“晚就诊者”？"),
    ("02", "病理生理：12 小时为什么不是绝对界限"),
    ("03", "关键证据：BRAVE-2、OAT、FAST-MI、影像学研究"),
    ("04", "2025–2026 最新证据：荟萃分析、LATE-MI、ISACS"),
    ("05", "国内外指南对比"),
    ("06", "该不该开：按时间与临床状态分层"),
    ("07", "什么时候开：时机选择与术前评估"),
    ("08", "术中要点、决策流程与总结"),
]
for i, (n, t) in enumerate(items):
    col, row = i % 2, i // 2
    x = Inches(0.8 + col * 6.1)
    y = Inches(1.6 + row * 1.25)
    box(s, x, y, Inches(5.8), Inches(1.0), fill=LIGHT)
    pill(s, x + Inches(0.15), y + Inches(0.18), Inches(0.9), Inches(0.64), n, RED if i % 2 else NAVY)
    text(s, x + Inches(1.2), y, Inches(4.5), Inches(1.0), t, size=19, bold=True,
         anchor=MSO_ANCHOR.MIDDLE)

# 3 临床问题
s = new_slide("临床问题：晚就诊 STEMI 并不少见", "症状发作至首次医疗接触 >12 小时，常被称为 late presenters / latecomers")
card(s, Inches(0.5), Inches(1.8), Inches(4.0), Inches(4.6), "有多常见", [
    ("FAST-MI 法国全国登记（2005/2010/2015）：6,273 例 STEMI 中 1,169 例（18.6%）在 12–48 h 就诊", {"bullet": True}),
    ("在院前急救体系薄弱、地域辽阔的地区比例更高", {"bullet": True}),
], color=NAVY)
card(s, Inches(4.67), Inches(1.8), Inches(4.0), Inches(4.6), "患者有什么特点", [
    ("年龄更大，女性、糖尿病、高血压更多", {"bullet": True}),
    ("既往卒中、心衰、肿瘤史更常见", {"bullet": True}),
    ("典型胸痛更少，心源性休克更多", {"bullet": True}),
    ("接受造影和 PCI 的比例更低", {"bullet": True}),
], color=TEAL)
card(s, Inches(8.83), Inches(1.8), Inches(4.0), Inches(4.6), "核心矛盾", [
    ("晚就诊者整体预后更差", {"bullet": True}),
    ("可挽救的心肌在减少，而操作风险（无复流、心肌内出血、出血并发症）依然存在", {"bullet": True}),
    ("直到 2026 年，针对这一人群的随机试验只有 3 项：BRAVE-2、OAT、LATE-MI", {"bullet": True, "bold": True, "color": RED}),
], color=RED)
source(s, "Bouisset F, et al. JACC 2021;78:1291-1305（FAST-MI）")

# 4 病理生理
s = new_slide("病理生理：为什么 12 小时不是一刀切", "12 h 界限主要来自溶栓时代（FTT 协作组）的数据，不能直接套用到直接 PCI")
card(s, Inches(0.5), Inches(1.8), Inches(6.0), Inches(4.7), "心肌坏死不一定在 12 h 内完成", [
    ("波阵面现象：坏死从心内膜向心外膜推进，速度因人而异", {"bullet": True}),
    ("侧支循环：Rentrop 分级高者，存活心肌可维持很久", {"bullet": True}),
    ("间断闭塞（stuttering）：血管反复开放与闭塞，“真正”的缺血时间短于症状时间", {"bullet": True}),
    ("缺血预适应：前驱心绞痛可延缓坏死", {"bullet": True}),
    ("症状发作时间常常估计不准", {"bullet": True}),
], color=TEAL)
card(s, Inches(6.83), Inches(1.8), Inches(6.0), Inches(4.7), "“开放血管假说”的晚期获益", [
    ("挽救残存的存活心肌（尤其是心外膜下心肌）", {"bullet": True}),
    ("改善左心室重构，减少室壁瘤和扩张", {"bullet": True}),
    ("提高电稳定性，减少室性心律失常", {"bullet": True}),
    ("为其他血管提供侧支储备", {"bullet": True}),
    ("但 OAT 证明：对于 3–28 天、稳定、无缺血的闭塞血管，“晚开”带来的这些理论获益并没有转化为临床获益", {"bullet": True, "bold": True, "color": RED}),
], color=NAVY)
source(s, "Trends Cardiovasc Med 2026（doi:10.1016/j.tcm.2026.05.002）：建议以“生物学分期”代替单纯时间切点")

# 5 获益与风险
s = new_slide("晚期开通的获益与风险", "决策的本质：剩余可挽救心肌 vs. 操作相关风险")
card(s, Inches(0.5), Inches(1.8), Inches(6.0), Inches(4.7), "支持“开”的因素", [
    ("持续或反复胸痛", {"bullet": True}),
    ("ST 段仍明显抬高，T 波尚未倒置，Q 波不明显", {"bullet": True}),
    ("血流动力学或电活动不稳定", {"bullet": True}),
    ("大面积心肌受累（前壁、近端病变）", {"bullet": True}),
    ("肌钙蛋白仍在上升（提示仍在坏死）", {"bullet": True}),
    ("影像学：水肿区大于坏死区，提示仍有存活心肌", {"bullet": True}),
], color=TEAL)
card(s, Inches(6.83), Inches(1.8), Inches(6.0), Inches(4.7), "提醒“慎开”的因素", [
    ("无症状、血流动力学稳定，症状已超过 48 h", {"bullet": True}),
    ("ST 段已回落，Q 波形成，T 波深倒", {"bullet": True}),
    ("透壁坏死、没有存活心肌", {"bullet": True}),
    ("晚期再灌注更容易出现无复流、微血管阻塞和心肌内出血", {"bullet": True}),
    ("出血风险高、肾功能差、高龄伴严重合并症", {"bullet": True}),
    ("OAT：晚期 PCI 有增加再梗死的趋势", {"bullet": True}),
], color=RED)

# 6 证据时间线
s = new_slide("证据全景：二十年的探索", "随机证据少，观察性证据多；2026 年新增两项重要研究")
events = [
    ("2005", "BRAVE-2", "12–48 h 稳定患者\n梗死面积缩小", TEAL),
    ("2006", "OAT", "3–28 天闭塞血管\nPCI 无获益", RED),
    ("2009", "BRAVE-2 长期", "4 年死亡率下降", TEAL),
    ("2018", "CMR 研究", "65% 晚就诊者仍有\n显著心肌挽救", TEAL),
    ("2021", "FAST-MI", "12–48 h 血运重建\n死亡风险降 35%", TEAL),
    ("2025–26", "荟萃分析", "12–72 h PCI\n死亡率降约 60%", TEAL),
    ("2026", "LATE-MI", "稳定者立即 PCI\n不优于延迟 PCI", AMBER),
    ("2026", "ISACS", "获益集中在\n25–48 h 窗口", AMBER),
]
line_y = Inches(3.55)
box(s, Inches(0.6), line_y, Inches(12.1), Inches(0.08), fill=GREY, shape=MSO_SHAPE.RECTANGLE)
for i, (yr, name, desc, col) in enumerate(events):
    x = Inches(0.6 + i * 1.53)
    dot = s.shapes.add_shape(MSO_SHAPE.OVAL, x + Inches(0.55), line_y - Inches(0.14), Inches(0.36), Inches(0.36))
    dot.fill.solid(); dot.fill.fore_color.rgb = col; dot.line.fill.background()
    up = i % 2 == 0
    yy = Inches(1.75) if up else Inches(4.05)
    box(s, x, yy, Inches(1.46), Inches(1.6), fill=LIGHT, line=col)
    text(s, x, yy + Inches(0.05), Inches(1.46), Inches(0.35), yr, size=13, color=col, bold=True, align=PP_ALIGN.CENTER)
    text(s, x, yy + Inches(0.38), Inches(1.46), Inches(0.4), name, size=13, bold=True, align=PP_ALIGN.CENTER)
    text(s, x, yy + Inches(0.78), Inches(1.46), Inches(0.8), desc, size=11, color=GREY, align=PP_ALIGN.CENTER)
text(s, Inches(0.6), Inches(5.95), Inches(12), Inches(0.5),
     "绿色：支持开通；红色：不支持常规开通；橙色：关于时机与人群选择的新证据", size=13, color=GREY)

# 7 BRAVE-2
s = new_slide("BRAVE-2：唯一针对 12–48 h 的“开 vs 不开”随机试验", "Schömig A, et al. JAMA 2005;293:2865-72；Ndrepepa G, et al. JAMA 2009;301:487-8")
table(s, Inches(0.5), Inches(1.8), Inches(6.3), Inches(3.4), [
    ["项目", "内容"],
    ["设计", "多中心、开放标签随机对照试验，2001–2004"],
    ["人群", "365 例，18–80 岁，12–48 h 就诊，无持续症状、无休克和严重心衰"],
    ["干预", "立即支架 + 阿昔单抗（182 例）vs 保守治疗（183 例）"],
    ["主要终点", "第 5–10 天 SPECT 测量的最终梗死面积"],
], col_w=[1.5, 4.8], font=13)
card(s, Inches(7.1), Inches(1.8), Inches(5.7), Inches(4.6), "主要结果", [
    ("梗死面积：8.0% vs 13.0%（中位数，P<0.001）", {"bullet": True, "bold": True}),
    ("24–48 h 亚组：6.5% vs 15.0%（P<0.001）", {"bullet": True}),
    ("12–24 h 亚组：9.0% vs 10.5%（P=0.06）", {"bullet": True}),
    ("30 天死亡/再梗/卒中：4.4% vs 6.6%（P=0.37，无统计学差异）", {"bullet": True}),
    ("4 年死亡率：11.1% vs 18.9%，校正 HR 0.55", {"bullet": True, "bold": True, "color": TEAL}),
], color=TEAL)
text(s, Inches(0.5), Inches(5.35), Inches(6.3), Inches(1.2), [
    ("局限：样本量小；硬终点不是主要终点；使用阿昔单抗和旧一代支架，与当前实践有差距。", {"color": GREY}),
], size=14)

# 8 OAT
s = new_slide("OAT：太晚、太稳定时，开了也没用", "Hochman JS, et al. NEJM 2006;355:2395-407")
table(s, Inches(0.5), Inches(1.8), Inches(6.3), Inches(3.4), [
    ["项目", "内容"],
    ["人群", "2,166 例，心梗后 3–28 天，梗死相关动脉完全闭塞，临床稳定"],
    ["高危标准", "LVEF <50% 或近端闭塞"],
    ["排除", "持续缺血、休克、NYHA III–IV 级、左主干或三支病变"],
    ["分组", "PCI + 最佳药物治疗 vs 单纯最佳药物治疗"],
], col_w=[1.5, 4.8], font=13)
card(s, Inches(7.1), Inches(1.8), Inches(5.7), Inches(4.6), "主要结果（4 年）", [
    ("死亡/再梗/IV 级心衰：17.2% vs 15.6%，HR 1.16（95% CI 0.92–1.45）", {"bullet": True, "bold": True}),
    ("死亡：9.1% vs 9.4%", {"bullet": True}),
    ("非致死性再梗死：6.9% vs 5.0%（P=0.08，PCI 组有增加趋势）", {"bullet": True}),
    ("1 年血管通畅率 83% vs 25%，但临床事件没有减少", {"bullet": True}),
    ("所有亚组（包括按时间分）均无交互作用", {"bullet": True}),
], color=RED)
text(s, Inches(0.5), Inches(5.35), Inches(6.3), Inches(1.2), [
    ("启示：这是指南中“>48 h 稳定、无症状者不常规开通”（III 类）的主要依据。OAT 的 PCI 中位时间约为第 8 天，与 BRAVE-2 不可直接比较。", {"color": DARK}),
], size=14)

# 9 FAST-MI
s = new_slide("FAST-MI 登记：真实世界中的 12–48 h 晚就诊者", "Bouisset F, et al. JACC 2021;78:1291-1305")
for i, (num, lab, col) in enumerate([
    ("1,077", "纳入分析的晚就诊者\n（729 例 48 h 内血运重建）", NAVY),
    ("2.1% vs 7.2%", "30 天全因死亡\n（P<0.001）", TEAL),
    ("30.4 vs 78.7", "长期死亡（每千人年）\n中位随访 58 个月", TEAL),
    ("HR 0.65", "多因素校正后死亡风险\n（95% CI 0.50–0.84）", RED),
]):
    x = Inches(0.5 + i * 3.1)
    box(s, x, Inches(1.85), Inches(2.9), Inches(2.2), fill=LIGHT, line=col)
    text(s, x, Inches(2.0), Inches(2.9), Inches(0.9), num, size=30, bold=True, color=col, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(s, x, Inches(2.95), Inches(2.9), Inches(1.0), lab, size=13, color=GREY, align=PP_ALIGN.CENTER)
text(s, Inches(0.5), Inches(4.35), Inches(12.3), Inches(2.2), [
    ("倾向性评分匹配（267 对）结果一致；再梗死也更少（5.4 vs 11.0 / 千人年）", {"bullet": True}),
    ("排除了溶栓患者和入院 48 h 内死亡者，以减少“存活偏倚”", {"bullet": True}),
    ("局限：观察性研究，存在选择偏倚——医生更倾向给状态好的患者做 PCI", {"bullet": True, "color": GREY}),
    ("这是 2023 ESC 指南推荐 12–48 h 常规直接 PCI（IIa, B）的重要依据之一", {"bullet": True, "bold": True, "color": NAVY}),
], size=16)

# 10 CMR
s = new_slide("影像学证据：晚开仍可挽救相当一部分心肌", "Nepper-Christensen L, et al. Circ Cardiovasc Interv 2018;11:e006842（DANAMI-3 CMR 亚组）")
table(s, Inches(0.5), Inches(1.8), Inches(7.2), Inches(2.6), [
    ["CMR 指标", "<12 h（n=807）", "12–72 h 有持续缺血（n=58）"],
    ["危险区面积", "33%", "34%"],
    ["最终梗死面积（中位）", "11%", "13%（P=0.037）"],
    ["心肌挽救指数（中位）", "0.65", "0.58（P=0.021）"],
    ["挽救指数 ≥0.50 的比例", "—", ("65%", TEAL)],
], col_w=[2.6, 2.0, 2.6], font=14)
card(s, Inches(8.0), Inches(1.8), Inches(4.8), Inches(4.6), "解读", [
    ("晚就诊者挽救的心肌确实更少，但差距并不大", {"bullet": True}),
    ("约 2/3 的晚就诊者仍获得显著心肌挽救", {"bullet": True, "bold": True}),
    ("从发病到 PCI 的时间预测心肌挽救的能力很差（AUC 0.58，R² <0.05）", {"bullet": True}),
    ("提示：应按“还剩多少存活心肌”来决策，而不是只看时钟", {"bullet": True, "color": RED, "bold": True}),
], color=TEAL)
text(s, Inches(0.5), Inches(4.6), Inches(7.2), Inches(1.8), [
    ("注意：这项研究的晚就诊者都有持续缺血的表现，结论不能直接推广到完全稳定、无症状的患者。", {"color": GREY}),
], size=14)

# 11 荟萃分析
s = new_slide("2025–2026 系统评价与荟萃分析", "观察性研究为主，方向一致：12–72 h 内 PCI 与更好预后相关")
card(s, Inches(0.5), Inches(1.8), Inches(6.0), Inches(4.7), "Mussolini 等，JACC 2026（ACC.26 摘要）", [
    ("9 项研究（1 项 RCT + 8 项观察性），7,237 例，12–72 h 就诊", {"bullet": True}),
    ("住院死亡：1.93% vs 5.21%，RR 0.39", {"bullet": True}),
    ("30 天死亡：2.94% vs 7.37%，RR 0.42", {"bullet": True}),
    ("12 个月死亡：5.21% vs 13.96%，RR 0.39", {"bullet": True, "bold": True}),
    ("再梗死 RR 0.65；心衰 RR 0.58", {"bullet": True}),
    ("卒中和出血无显著差异", {"bullet": True}),
], color=NAVY)
card(s, Inches(6.83), Inches(1.8), Inches(6.0), Inches(4.7), "Fatima 等，Cardiovasc Revasc Med 2025", [
    ("8 项研究，聚焦稳定的 12–72 h 晚就诊者", {"bullet": True}),
    ("PCI vs 药物（5 项，3,820 例）：PCI 降低全因死亡、再梗死和梗死面积，提高心肌挽救指数", {"bullet": True}),
    ("紧急 vs 非紧急 PCI（3 项，1,270 例）：临床结局没有显著差异", {"bullet": True, "bold": True, "color": RED}),
    ("证据质量中等，偏倚风险中至高", {"bullet": True, "color": GREY}),
], color=TEAL)
source(s, "J Am Coll Cardiol 2026; doi:10.1016/j.jacc.2026.02.1035 · PMID 39779400")

# 12 LATE-MI
s = new_slide("LATE-MI（ESC 2026）：稳定者需要“立即”开吗？", "Nepper-Christensen L, Lønborg J, Engstrøm T, et al. 丹麦 Rigshospitalet；NCT02445885")
table(s, Inches(0.5), Inches(1.8), Inches(6.5), Inches(3.3), [
    ["项目", "内容"],
    ["人群", "症状发作 >12 h、临床稳定的 STEMI"],
    ["分组", "立即 PCI（按直接 PCI 流程）vs 亚急性/延迟 PCI（72 h 内）"],
    ["样本量", "约 150 例（登记数据）"],
    ["主要终点", "CMR 急性心肌挽救指数"],
    ["次要终点", "最终梗死面积、MVO、LVEF、死亡/心衰/恶性心律失常"],
], col_w=[1.5, 5.0], font=13)
card(s, Inches(7.3), Inches(1.8), Inches(5.5), Inches(4.6), "结果（大会报告）", [
    ("立即 PCI 未能改善：", {"bold": True}),
    ("心肌挽救指数", {"bullet": True}),
    ("最终梗死面积", {"bullet": True}),
    ("左心室功能", {"bullet": True}),
    ("临床事件", {"bullet": True}),
    ("要回答的问题不是“开不开”，而是对稳定者“要不要半夜急诊开”", {"bold": True, "color": RED}),
], color=AMBER)
text(s, Inches(0.5), Inches(5.3), Inches(6.5), Inches(1.3), [
    ("注意：样本量小、以影像学为主要终点；截至本幻灯片制作时，全文和详细数据尚待正式发表，请以发表版本为准。", {"color": GREY}),
], size=13)

# 13 ISACS
s = new_slide("ISACS 登记（EHJ 2026）：获益随时间变化", "Eur Heart J 2026; doi:10.1093/eurheartj/ehag785.1789（ESC 2026 摘要）")
text(s, Inches(0.5), Inches(1.8), Inches(6.2), Inches(2.0), [
    ("2,730 例临床稳定、12–48 h 就诊的 STEMI 患者", {"bullet": True}),
    ("排除休克、心脏骤停、严重心衰、持续心绞痛", {"bullet": True}),
    ("方法：逆概率加权（IPTW）+ 倾向性匹配", {"bullet": True}),
    ("主要终点：30 天死亡", {"bullet": True}),
], size=16)
table(s, Inches(0.5), Inches(4.0), Inches(6.2), Inches(1.9), [
    ["分组", "30 天死亡 / 效应"],
    ["整体：PCI vs 药物", "3.1% vs 5.5%，OR 0.54（0.36–0.84）"],
    ["PCI 在 12–24 h 完成", "风险校正死亡率 6.3%"],
    ["PCI 在 25–48 h 完成", ("1.8%；vs 药物 OR 0.30（0.16–0.58）", TEAL)],
], col_w=[2.4, 3.8], font=13)
card(s, Inches(7.0), Inches(1.8), Inches(5.8), Inches(4.6), "怎么理解", [
    ("交互作用 P=0.001：获益大小与时间相关", {"bullet": True}),
    ("对稳定者，“越早越好”并不总成立", {"bullet": True}),
    ("可能的原因：12–24 h 组包含更多隐匿的高危患者；或者早期干预更容易出现再灌注损伤", {"bullet": True}),
    ("与 LATE-MI 一起指向同一结论：稳定的晚就诊者应该开，但不必急诊开", {"bullet": True, "bold": True, "color": RED}),
    ("属于观察性摘要，结论需要谨慎看待", {"bullet": True, "color": GREY}),
], color=AMBER)

# 14 指南对比
s = new_slide("指南对比：>12 小时 STEMI 的再灌注推荐", "推荐类别 / 证据等级")
table(s, Inches(0.4), Inches(1.7), Inches(12.5), Inches(4.7), [
    ["临床情境", "2023 ESC ACS", "2025 ACC/AHA ACS", "2019 中国 STEMI", "2025 中国 PCI"],
    [">12 h，持续缺血 / 血流动力学不稳定 / 致命性心律失常", ("推荐 PPCI（I, C）", TEAL), ("休克/不稳定：PPCI（I）；>24 h 仍缺血：合理（IIa）", TEAL), ("推荐 PPCI（I, B）", TEAL), ("机械循环支持下考虑 PPCI（IIa, A）", TEAL)],
    ["12–24 h，稳定、无症状", ("考虑常规 PPCI（IIa, B）", TEAL), ("PPCI 合理（IIa）", TEAL), ("考虑 PCI（IIa, B）", TEAL), ("可常规造影，必要时开通（IIa, A）", TEAL)],
    ["24–48 h，稳定、无症状", ("考虑常规 PPCI（IIa, B）", TEAL), ("证据不足，未明确推荐", AMBER), ("考虑 PCI（IIa, B）", TEAL), ("原则上不推荐 PPCI（IIb, A）", AMBER)],
    [">48 h，稳定、无症状、IRA 闭塞", ("不推荐常规 PCI（III, A）", RED), ("梗死已完成的闭塞 IRA 不推荐 PCI（III：无获益）", RED), ("不推荐直接 PCI（III, A）", RED), ("不推荐，可能有害（III-Harm, B-R）", RED)],
    [">12 h 能否溶栓", ("任何情况都优先 PPCI，不推荐溶栓", DARK), ("溶栓仅限 12 h 内", DARK), ("12–24 h 仍有缺血且无 PCI 条件：考虑溶栓（IIa, C）", DARK), ("—", DARK)],
], col_w=[3.1, 2.25, 2.55, 2.3, 2.3], font=12)
source(s, "注：ACC/AHA 2025 认为 PPCI 获益可延续至约 24 h，24–48 h 无症状者“验证较少”。ACC/AHA 列仅标推荐类别。中国 PCI 指南英文版发表于 Cardiology Discovery 2026。")

# 15 该开的情况
s = new_slide("该开：任何时间都应立即开通的情况", "各大指南一致：出现以下任一情况，症状时长不是限制条件")
for i, (t, d, c) in enumerate([
    ("持续或反复缺血", "胸痛不缓解，ST 段动态变化\n（I 类）", RED),
    ("心源性休克", "立即 PPCI，不论时间（I 类）\n必要时机械循环支持", RED),
    ("急性严重心衰", "缺血导致的肺水肿、\n低心排", RED),
    ("致命性心律失常", "反复 VT/VF、\n心脏骤停复苏后", RED),
    ("机械并发症", "室间隔穿孔、乳头肌断裂、\n游离壁破裂：首选外科", AMBER),
]):
    x = Inches(0.5 + i * 2.5)
    box(s, x, Inches(1.85), Inches(2.3), Inches(2.7), fill=WHITE, line=c)
    pill(s, x + Inches(0.1), Inches(1.95), Inches(2.1), Inches(0.75), t, c)
    text(s, x + Inches(0.05), Inches(2.85), Inches(2.2), Inches(1.6), d, size=14, align=PP_ALIGN.CENTER)
text(s, Inches(0.5), Inches(4.85), Inches(12.3), Inches(1.7), [
    ("2025 ACC/AHA：心源性休克或血流动力学不稳定者应尽快 PPCI，理想 ≤90 min；多支病变只处理罪犯血管（CULPRIT-SHOCK）", {"bullet": True}),
    ("STEMI 合并休克：在选择性患者中可考虑微轴流泵（DanGer Shock，NEJM 2024）；不推荐常规使用 IABP", {"bullet": True}),
    ("2025 中国 PCI 指南：此类极高危患者应考虑在机械循环支持下行直接 PCI（IIa, A）", {"bullet": True}),
], size=15)

# 16 时间分层
s = new_slide("该不该开：按时间与临床状态分层", "适用于临床稳定、无持续缺血的晚就诊者（不稳定者见上一页）")
cols = [
    ("12–24 h", TEAL, "建议开", ["证据最一致：所有指南均为 IIa", "BRAVE-2、FAST-MI、荟萃分析均支持", "目标：在本次住院中完成造影 ± PCI"]),
    ("24–48 h", AMBER, "倾向开，个体化", ["ESC/中国 2019：IIa", "中国 2025：IIb；ACC/AHA：证据不足", "ISACS：获益可能集中在这一窗口", "看 ST 段、Q 波、症状和心肌存活情况"]),
    (">48 h", RED, "原则上不急诊开", ["稳定、无症状、IRA 闭塞：III 类", "依据：OAT", "先做缺血或存活心肌评估", "有缺血证据时再择期血运重建"]),
]
for i, (t, c, verdict, pts) in enumerate(cols):
    x = Inches(0.5 + i * 4.17)
    box(s, x, Inches(1.85), Inches(3.95), Inches(4.65), fill=WHITE, line=c)
    pill(s, x, Inches(1.85), Inches(3.95), Inches(0.7), t, c)
    text(s, x, Inches(2.65), Inches(3.95), Inches(0.6), verdict, size=22, bold=True, color=c, align=PP_ALIGN.CENTER)
    text(s, x + Inches(0.15), Inches(3.35), Inches(3.7), Inches(3.0), [(p, {"bullet": True}) for p in pts], size=15)

# 17 何时开
s = new_slide("什么时候开：时机选择", "核心原则：不稳定者立即开；稳定者“要开，但不必急诊开”")
table(s, Inches(0.5), Inches(1.75), Inches(12.3), Inches(3.3), [
    ["临床情境", "推荐时机", "依据"],
    ["任何时间 + 持续缺血 / 休克 / 心衰 / 恶性心律失常", ("立即，按直接 PCI 流程（诊断到导丝通过 ≤90–120 min）", RED), "各指南 I 类"],
    ["12–48 h，稳定、无症状", ("本次住院内尽早完成（可安排在 24–72 h 内的常规工作时段）", TEAL), "LATE-MI；紧急 vs 非紧急 PCI 结局相似"],
    [">48 h，稳定、无症状", ("不急诊开；先做无创缺血或存活心肌评估，有指征时择期血运重建", AMBER), "OAT；III 类"],
    ["等待期间再次出现缺血或不稳定", ("随时转为立即 PCI", RED), "补救原则"],
], col_w=[3.6, 5.4, 3.3], font=13)
text(s, Inches(0.5), Inches(5.25), Inches(12.3), Inches(1.4), [
    (">12 h 不推荐溶栓代替 PCI（中国 2019 指南：12–24 h 仍有缺血且无 PCI 条件时可考虑溶栓，IIa, C）", {"bullet": True}),
    ("延迟策略期间：双联抗血小板、抗凝、β 受体阻滞剂、他汀，持续心电监护", {"bullet": True}),
], size=15)

# 18 术前评估
s = new_slide("个体化评估：用“生物学时钟”代替“时间时钟”", "帮助判断是否还有可挽救的心肌")
for i, (t, pts, c) in enumerate([
    ("症状与体征", ["胸痛是否持续、反复", "Killip 分级、血压、心率", "心律失常情况"], NAVY),
    ("心电图", ["ST 段是否仍明显抬高", "T 波直立 vs 深倒（急性程度）", "Q 波深度与范围", "Anderson-Wilkins 急性程度评分"], TEAL),
    ("生物标志物", ["肌钙蛋白仍在上升还是已过峰值", "BNP / NT-proBNP"], AMBER),
    ("影像学", ["超声：室壁运动、机械并发症、LVEF", "CMR：水肿（危险区）vs 延迟强化（坏死）", "存活心肌评估（多巴酚丁胺超声、核素）"], RED),
]):
    x = Inches(0.5 + i * 3.1)
    card(s, x, Inches(1.85), Inches(2.95), Inches(3.7), t, [(p, {"bullet": True}) for p in pts], color=c, size=14)
text(s, Inches(0.5), Inches(5.75), Inches(12.3), Inches(0.9), [
    ("同时评估操作风险：年龄、出血风险（PRECISE-DAPT / ARC-HBR）、肾功能、合并症和患者意愿", {"bullet": True, "bold": True}),
], size=15)

# 19 术中要点
s = new_slide("术中要点：晚期开通更要防“无复流”", "晚期再灌注时微血管阻塞（MVO）和心肌内出血更常见")
card(s, Inches(0.5), Inches(1.8), Inches(6.0), Inches(4.7), "操作策略", [
    ("首选桡动脉入路（I, A）", {"bullet": True}),
    ("新一代药物洗脱支架；用 IVUS/OCT 优化（IIa）", {"bullet": True}),
    ("避免支架过度后扩张；血栓负荷大时考虑选择性血栓抽吸（不推荐常规）", {"bullet": True}),
    ("休克或不稳定者只处理罪犯血管，非罪犯血管择期处理", {"bullet": True}),
    ("稳定的多支病变：推荐完全血运重建（可分次）", {"bullet": True}),
], color=NAVY)
card(s, Inches(6.83), Inches(1.8), Inches(6.0), Inches(4.7), "无复流的预防与处理", [
    ("冠脉内给药：替罗非班、腺苷、硝普钠、钙通道阻滞剂、硝酸酯（IIa, B）", {"bullet": True}),
    ("尽量缩短导丝通过到恢复血流的时间", {"bullet": True}),
    ("严重无复流伴血流动力学不稳定时，可用 IABP 等循环支持", {"bullet": True}),
    ("抗栓：阿司匹林 + 替格瑞洛（或氯吡格雷）；术中肠外抗凝", {"bullet": True}),
    ("术后 CMR 可用于评估 MVO 和预后", {"bullet": True}),
], color=TEAL)
source(s, "2019 中国 STEMI 指南；2025 中国 PCI 指南；2023 ESC ACS 指南")

# 20 流程图 + 总结
s = new_slide("决策流程与总结")
fx = Inches(0.4)
pill(s, fx, Inches(1.35), Inches(6.8), Inches(0.6), "STEMI，症状发作 >12 小时", NAVY)
down_arrow(s, fx + Inches(3.2), Inches(1.98), Inches(0.4), Inches(0.3))
pill(s, fx, Inches(2.3), Inches(6.8), Inches(0.6), "持续缺血 / 休克 / 心衰 / 恶性心律失常？", GREY)
text(s, fx, Inches(2.95), Inches(1.6), Inches(0.35), "是", size=14, bold=True, color=RED)
text(s, fx + Inches(3.7), Inches(2.95), Inches(1.6), Inches(0.35), "否（稳定）", size=14, bold=True, color=TEAL)
pill(s, fx, Inches(3.35), Inches(3.0), Inches(0.9), "立即 PPCI\n（不论时间）", RED)
for i, (lab, verdict, c) in enumerate([
    ("12–24 h", "开：住院内尽早", TEAL),
    ("24–48 h", "倾向开：个体化", AMBER),
    (">48 h", "先评估缺血/存活", RED),
]):
    y = Inches(3.35 + i * 1.05)
    pill(s, fx + Inches(3.3), y, Inches(1.2), Inches(0.9), lab, c)
    box(s, fx + Inches(4.6), y, Inches(2.2), Inches(0.9), fill=LIGHT, line=c)
    text(s, fx + Inches(4.6), y, Inches(2.2), Inches(0.9), verdict, size=13, bold=True, color=c,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
text(s, fx, Inches(4.45), Inches(3.0), Inches(1.9), [
    ("等待期间如再发缺血或不稳定，随时转为立即 PCI", {"color": RED}),
], size=13)

box(s, Inches(7.5), Inches(1.35), Inches(5.4), Inches(5.2), fill=LIGHT)
text(s, Inches(7.7), Inches(1.45), Inches(5.0), Inches(0.5), "Take-home messages", size=20, bold=True, color=NAVY)
text(s, Inches(7.7), Inches(2.0), Inches(5.0), Inches(4.5), [
    ("1. 12 h 不是绝对界限，“剩余存活心肌”比时钟更重要", {"after": 10}),
    ("2. 不稳定者：任何时间都应立即开通", {"after": 10, "bold": True, "color": RED}),
    ("3. 稳定的 12–48 h：开通获益证据一致（BRAVE-2、FAST-MI、荟萃分析）", {"after": 10}),
    ("4. 稳定者不必急诊开：LATE-MI 显示立即与延迟 PCI 结局相似", {"after": 10, "bold": True, "color": TEAL}),
    ("5. >48 h 稳定、无症状、IRA 闭塞：不常规开通（OAT）", {"after": 10}),
    ("6. >12 h 不用溶栓代替 PCI", {"after": 10}),
], size=15)

# 21 参考文献
s = new_slide("参考文献")
text(s, Inches(0.5), Inches(1.35), Inches(12.3), Inches(5.6), [
    "1. Byrne RA, et al. 2023 ESC Guidelines for the management of acute coronary syndromes. Eur Heart J 2023;44:3720-3826.",
    "2. Rao SV, et al. 2025 ACC/AHA/ACEP/NAEMSP/SCAI Guideline for the Management of Patients With Acute Coronary Syndromes. Circulation 2025; doi:10.1161/CIR.0000000000001309.",
    "3. 中华医学会心血管病学分会. 急性ST段抬高型心肌梗死诊断和治疗指南（2019）. 中华心血管病杂志 2019;47:766-783.",
    "4. 中华医学会心血管病学分会. 中国经皮冠状动脉介入治疗指南（2025）. 中华心血管病杂志 2025;53:717-745；英文版 Cardiology Discovery 2026.",
    "5. Schömig A, et al. Mechanical reperfusion in patients with AMI presenting >12 hours from symptom onset (BRAVE-2). JAMA 2005;293:2865-72.",
    "6. Ndrepepa G, et al. Mechanical reperfusion and long-term mortality in patients with AMI presenting 12 to 48 hours from onset of symptoms. JAMA 2009;301:487-8.",
    "7. Hochman JS, et al. Coronary intervention for persistent occlusion after myocardial infarction (OAT). N Engl J Med 2006;355:2395-407.",
    "8. Bouisset F, et al. Percutaneous myocardial revascularization in late-presenting patients with STEMI (FAST-MI). J Am Coll Cardiol 2021;78:1291-1305.",
    "9. Nepper-Christensen L, et al. Benefit from reperfusion with primary PCI beyond 12 hours of symptom duration in STEMI. Circ Cardiovasc Interv 2018;11:e006842.",
    "10. Fatima S, et al. PCI for stable late STEMI with symptom onset between 12 and 72 h: a systematic review. Cardiovasc Revasc Med 2025;78:1-9.",
    "11. Mussolini MCB, et al. PCI in late-presenting STEMI: a systematic review and meta-analysis. J Am Coll Cardiol 2026 (ACC.26 abstract); doi:10.1016/j.jacc.2026.02.1035.",
    "12. Nepper-Christensen L, et al. Immediate versus deferred PCI in clinically stable patients with STEMI presenting beyond 12 hours (LATE-MI). ESC Congress 2026, Late-Breaking Clinical Trial; NCT02445885.",
    "13. Time-dependent survival benefit of late PCI in stable STEMI: a propensity-weighted analysis of the 12–48 hour window (ISACS). Eur Heart J 2026; doi:10.1093/eurheartj/ehag785.1789.",
    "14. Late-presenting STEMI: is it just a matter of time? From temporal cutoffs to biological staging. Trends Cardiovasc Med 2026; doi:10.1016/j.tcm.2026.05.002.",
    "15. Møller JE, et al. Microaxial flow pump or standard care in infarct-related cardiogenic shock (DanGer Shock). N Engl J Med 2024;390:1382-93.",
], size=11, spacing=1.05)

assert len(prs.slides._sldIdLst) == TOTAL, len(prs.slides._sldIdLst)
prs.save("STEMI_发病超过12小时_再灌注决策.pptx")
print("saved", TOTAL)
