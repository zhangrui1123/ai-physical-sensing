# -*- coding: utf-8 -*-
"""AI 物理感知平台 PPT：基于华为浅色 16:9 模板母版。"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree

TEMPLATE = r"g:\My Drive\Documents\AI物理感知\PPT模板-浅色版16-9.pptx"
OUT = r"g:\My Drive\Documents\AI物理感知\AI物理感知平台-华为汇报.pptx"

BG = RGBColor(0xFF, 0xFF, 0xFF)
CARD = RGBColor(0xF7, 0xF7, 0xF7)
STROKE = RGBColor(0xDD, 0xDD, 0xDD)
RED = RGBColor(0xC7, 0x00, 0x0B)
NAVY = RGBColor(0x1A, 0x1A, 0x1A)
BODY = RGBColor(0x55, 0x57, 0x57)
MUTED = RGBColor(0x89, 0x89, 0x89)
CYAN = RGBColor(0x30, 0xB5, 0xC5)
GREEN = RGBColor(0x62, 0xB2, 0x30)
ORANGE = RGBColor(0xED, 0x6D, 0x00)
INDIGO = RGBColor(0x00, 0x74, 0xCC)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
FONT = "Microsoft YaHei"
TOTAL = 9


def set_run(run, size, color, bold=False, name=FONT):
    run.font.size = Pt(size)
    run.font.color.rgb = color
    run.font.bold = bold
    run.font.name = name
    rPr = run._r.get_or_add_rPr()
    for tag in ("ea", "cs"):
        el = rPr.find(qn(f"a:{tag}"))
        if el is None:
            el = etree.SubElement(rPr, qn(f"a:{tag}"))
        el.set("typeface", name)


def box(slide, l, t, w, h, fill, line=None):
    sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, t, w, h)
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
        sh.line.width = Pt(0.75)
    try:
        sh.adjustments[0] = 0.05
    except Exception:
        pass
    return sh


def rect(slide, l, t, w, h, fill):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    sh.line.fill.background()
    return sh


def put(slide, l, t, w, h, text, size, color, bold=False, align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP):
    sh = slide.shapes.add_textbox(l, t, w, h)
    tf = sh.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    try:
        tf._txBody.bodyPr.set(
            "anchor",
            {MSO_ANCHOR.TOP: "t", MSO_ANCHOR.MIDDLE: "ctr", MSO_ANCHOR.BOTTOM: "b"}[valign],
        )
    except Exception:
        pass
    for i, line in enumerate(str(text).split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(2)
        run = p.add_run()
        run.text = line
        set_run(run, size, color, bold)
    return tf


def delete_all_slides(prs):
    sldIdLst = prs.slides._sldIdLst
    for sldId in list(sldIdLst):
        rId = sldId.get(qn("r:id"))
        prs.part.drop_rel(rId)
        sldIdLst.remove(sldId)


def blank_slide(prs):
    layout = min(prs.slide_layouts, key=lambda L: len(L.placeholders))
    s = prs.slides.add_slide(layout)
    for sh in list(s.placeholders):
        sp = sh._element
        sp.getparent().remove(sp)
    rect(s, 0, 0, prs.slide_width, prs.slide_height, BG)
    return s


def footer(slide, page, section=""):
    rect(slide, 0, Inches(7.15), Inches(13.333), Inches(0.35), RGBColor(0xF2, 0xF2, 0xF2))
    put(
        slide, Inches(0.5), Inches(7.15), Inches(9.0), Inches(0.35),
        f"华为  ·  AI 物理感知平台  ·  {section}    |    Security Level: Internal",
        9, MUTED, valign=MSO_ANCHOR.MIDDLE,
    )
    put(
        slide, Inches(11.2), Inches(7.15), Inches(1.7), Inches(0.35),
        f"{page}  /  {TOTAL}", 9, MUTED, align=PP_ALIGN.RIGHT, valign=MSO_ANCHOR.MIDDLE,
    )


def header(slide, kicker, title, subtitle=None):
    put(slide, Inches(0.5), Inches(0.22), Inches(12.2), Inches(0.28), kicker, 11, RED, True)
    put(slide, Inches(0.5), Inches(0.48), Inches(12.2), Inches(0.48), title, 24, NAVY, True)
    rect(slide, Inches(0.5), Inches(1.0), Inches(1.2), Inches(0.05), RED)
    if subtitle:
        put(slide, Inches(0.5), Inches(1.12), Inches(12.2), Inches(0.32), subtitle, 12, MUTED)


def so_what(slide, text):
    box(slide, Inches(0.5), Inches(6.35), Inches(12.3), Inches(0.7), RGBColor(0xFF, 0xF0, 0xF0), RED)
    rect(slide, Inches(0.5), Inches(6.35), Inches(0.08), Inches(0.7), RED)
    put(slide, Inches(0.75), Inches(6.35), Inches(11.9), Inches(0.7),
        "所以  ·  " + text, 13, NAVY, True, valign=MSO_ANCHOR.MIDDLE)


def build():
    prs = Presentation(TEMPLATE)
    delete_all_slides(prs)

    # 1 封面
    s = blank_slide(prs)
    rect(s, 0, 0, Inches(0.12), prs.slide_height, RED)
    put(s, Inches(0.7), Inches(1.25), Inches(12), Inches(0.32),
        "华为内部立项  ·  痛点驱动：多模态 · 多设备 · 多场景", 14, RED, True)
    put(s, Inches(0.7), Inches(1.7), Inches(12), Inches(0.7), "AI 物理感知平台", 40, NAVY, True)
    put(s, Inches(0.7), Inches(2.5), Inches(12), Inches(0.4),
        "接入规格 → 共享特征 → 预测 / 分类双接口，不做场景应用", 18, BODY)
    box(s, Inches(0.7), Inches(3.1), Inches(11.9), Inches(1.2), CARD, STROKE)
    put(s, Inches(0.95), Inches(3.25), Inches(11.4), Inches(0.95),
        "三大痛点：多模态融合不足、多设备自适应困难、多任务多场景泛化弱。\n"
        "平台对策：统一 patch token 特征 + 统一推理接口；底座可换，接口不变。",
        15, BODY)
    for i, (k, v) in enumerate([
        ("痛点", "多模态 · 多设备 · 多场景"),
        ("接口", "接入规格 + 预测 + 分类"),
        ("演进", "零样本基线 → 千人千面 → Token 嵌入大模型"),
    ]):
        left = Inches(0.7 + i * 4.05)
        box(s, left, Inches(4.55), Inches(3.85), Inches(1.7), WHITE, STROKE)
        rect(s, left, Inches(4.55), Inches(3.85), Inches(0.08), RED)
        put(s, left + Inches(0.2), Inches(4.75), Inches(3.45), Inches(0.35), k, 13, RED, True)
        put(s, left + Inches(0.2), Inches(5.2), Inches(3.45), Inches(0.85), v, 15, NAVY)
    put(s, Inches(0.7), Inches(6.5), Inches(8.5), Inches(0.3), "来源：TimesFM / UniTS / MOMENT / LIMU-BERT-X 公开资料  ·  9 页平台版", 12, MUTED)
    put(s, Inches(10.3), Inches(6.5), Inches(2.5), Inches(0.3), f"1  /  {TOTAL}", 12, MUTED, align=PP_ALIGN.RIGHT)
    put(s, Inches(0.7), Inches(6.9), Inches(6), Inches(0.25), "Security Level: Internal", 10, MUTED)

    # 2 业务痛点
    s = blank_slide(prs)
    header(s, "1  业务痛点", "三个痛点，一个根因：缺统一特征与统一接口",
           "来自一线业务复盘；每个痛点都对应后文一类模型能力。")
    pains = [
        ("痛点一\n多模态融合不足", "异构模态缺乏统一对齐。\nIMU / EMG / TP 各自建模：量纲、采样率、时钟不一致；\n跨模态信息互相看不见，融合靠人工拼特征。", RED),
        ("痛点二\n多设备自适应困难", "基于单设备独立优化。\n每款设备单独采数据、单独调参；\n换机型、换佩戴位置就要重新标定，成本线性增长。", ORANGE),
        ("痛点三\n多任务多场景泛化弱", "算法针对特定场景设计。\n一个场景一个模型；新场景重新标注、重新训练；\n无法零样本启动，长尾场景覆盖不了。", CYAN),
    ]
    for i, (t, d, c) in enumerate(pains):
        left = Inches(0.45 + i * 3.2)
        box(s, left, Inches(1.55), Inches(3.05), Inches(3.9), WHITE, STROKE)
        rect(s, left, Inches(1.55), Inches(3.05), Inches(0.08), c)
        put(s, left + Inches(0.15), Inches(1.8), Inches(2.75), Inches(0.75), t, 15, c, True)
        put(s, left + Inches(0.15), Inches(2.65), Inches(2.75), Inches(2.6), d, 12, BODY)
    box(s, Inches(0.5), Inches(5.6), Inches(12.3), Inches(0.9), CARD, STROKE)
    put(s, Inches(0.7), Inches(5.75), Inches(11.9), Inches(0.6),
        "根因：没有「统一特征 + 统一接口」的物理时序底座——这正是平台的定位。",
        14, RED, True)
    footer(s, 2, "业务痛点")

    # 3 模型盘点·通用时序
    s = blank_slide(prs)
    header(s, "2  模型盘点 · 通用时序", "通用时序基础模型：预测强，分类只有两家原生支持",
           "口径：预训练目标决定原生能力；TimesFM / Chronos / Moirai 都只做预测。")
    cols = [(0.5, 2.5), (3.05, 2.5), (5.6, 3.6), (9.25, 1.4), (10.7, 2.1)]
    for (cx, cw), htxt in zip(cols, ["模型 / 作者单位", "预训练目标", "原生能力", "分类", "许可"]):
        put(s, Inches(cx + 0.12), Inches(1.42), Inches(cw - 0.2), Inches(0.32), htxt, 12, MUTED, True)
    rows = [
        ("TimesFM-3\nGoogle Research\nICML'24 / 3.0 2026", "decoder-only\n分位数回归", "多变量预测 + 协变量；\nCPM 一次前向出 9 分位", "✗ 规则/加头", "3.0 权重 NC\n≤2.5 Apache"),
        ("Chronos\nAmazon Science\nICML'24", "离散 token\nT5 式 NTP", "单变量概率预测", "✗", "Apache-2.0"),
        ("Moirai\nSalesforce AI Research\nICML'24", "masked encoder\nany-variate", "多变量概率预测", "✗", "代码 Apache\n权重 CC-BY-NC"),
        ("UniTS\nHarvard + MIT Lincoln Lab\nNeurIPS'24", "任务 token\nGEN/CLS 共享参数", "预测+分类+填补+异常一体；\n38 个数据集对比领先", "✓ 原生", "MIT"),
        ("MOMENT\nCMU Auton Lab\nICML'24", "掩码 patch 重建\nT5 encoder", "预测/分类/异常/填补；\n线性探针即用", "✓ 探针", "MIT"),
    ]
    for i, row in enumerate(rows):
        top = Inches(1.82 + i * 0.9)
        box(s, Inches(0.5), top, Inches(12.3), Inches(0.82), WHITE if i % 2 == 0 else CARD, STROKE)
        for j, ((cx, cw), cell) in enumerate(zip(cols, row)):
            if j == 0:
                put(s, Inches(cx + 0.12), top + Inches(0.04), Inches(cw - 0.2), Inches(0.76), cell, 11, RED, True)
            elif j == 3:
                put(s, Inches(cx + 0.12), top + Inches(0.06), Inches(cw - 0.2), Inches(0.72), cell, 12,
                    GREEN if cell.startswith("✓") else MUTED, True)
            else:
                put(s, Inches(cx + 0.12), top + Inches(0.06), Inches(cw - 0.2), Inches(0.72), cell, 11, BODY)
    so_what(s, "预测底座选 TimesFM-3（多变量+协变量最强）；原生分类只有 UniTS / MOMENT——这正是双头设计的来源。")
    footer(s, 3, "模型盘点")

    # 4 模型盘点·物理传感器
    s = blank_slide(prs)
    header(s, "3  模型盘点 · 物理传感器", "IMU / EMG 专用模型：提供领域先验，不当平台底座",
           "传感器模型规模小、端侧导向；与通用底座互补，不替代。")
    cols = [(0.5, 2.6), (3.15, 3.3), (6.5, 4.4), (10.95, 1.85)]
    for (cx, cw), htxt in zip(cols, ["模型 / 作者单位", "数据 / 预训练", "能力亮点", "部署 / 许可"]):
        put(s, Inches(cx + 0.12), Inches(1.42), Inches(cw - 0.2), Inches(0.32), htxt, 12, MUTED, True)
    rows = [
        ("LIMU-BERT\n南洋理工+阿里巴巴\nSenSys'21", "IMU 三轴\n掩码预训练", "端侧 HAR 开山作；小样本迁移到手机 / 腕戴", "手机可部署"),
        ("LIMU-BERT-X\n港科大+阿里巴巴\nMobiCom'25", "143 万小时 · 6 万人\n1.1K 种机型", "真实设备大规模泛化；端侧 HAR SOTA；外卖配送全国部署", "端侧"),
        ("TartanIMU\nCMU AirLab\nCVPR'25", "跨机器人 IMU\n位姿基础模型", "LoRA 仅 1.1M 参数即适配新机；200 FPS 在线推理", "机器人"),
        ("PRIMUS\nNokia Bell Labs\n+华盛顿大学 ICASSP'25", "IMU 多模态\n自监督对齐", "对齐预训练，提升下游 HAR / 健康任务", "研究"),
        ("Babel\n微软研究院+威斯康星\n麦迪逊+港科大 SenSys'25", "6 模态对齐\nIMU 塔 = LIMU-BERT", "多模态人体感知特征", "研究"),
        ("Meta EMG\nMeta\nNature'25", "sEMG 腕带\n跨用户泛化", "0.88 手势/s；手写 20.9 WPM；个性化再 +16%", "腕带\nCC-BY-NC"),
    ]
    for i, row in enumerate(rows):
        top = Inches(1.8 + i * 0.75)
        box(s, Inches(0.5), top, Inches(12.3), Inches(0.68), WHITE if i % 2 == 0 else CARD, STROKE)
        for j, ((cx, cw), cell) in enumerate(zip(cols, row)):
            if j == 0:
                put(s, Inches(cx + 0.12), top + Inches(0.03), Inches(cw - 0.2), Inches(0.64), cell, 10, RED, True)
            else:
                put(s, Inches(cx + 0.12), top + Inches(0.04), Inches(cw - 0.2), Inches(0.6), cell, 11, BODY)
    so_what(s, "平台用 LIMU-BERT-X 初始化 IMU 变元、用 Meta EMG 作 EMG 对照；底座仍是通用模型。")
    footer(s, 4, "模型盘点")

    # 5 TOP3 研究团队
    s = blank_slide(prs)
    header(s, "4  TOP3 研究团队", "物理感知方向最值得对标的三支队伍",
           "选型口径：有基础模型、有真实数据、有产业部署。")
    teams = [
        ("Google Research", "TimesFM 系列（1.0 → 3.0）。\n通用时序预测标杆：fev-bench / TIME / GIFT-Eval 三榜第一。\n平台预测底座的来源。", RED),
        ("Mo Li 团队\n港科大/南洋理工 + 阿里巴巴", "LIMU-BERT（SenSys'21 最佳论文提名）→ LIMU-BERT-X（MobiCom'25）→ Babel（SenSys'25）。\n143 万小时真实数据；外卖配送全国部署。\n传感器基础模型从论文走到产业。", ORANGE),
        ("Meta Reality Labs", "sEMG 神经腕带（Nature'25）。\n跨用户泛化：0.88 手势/s、手写 20.9 WPM。\n已随 Ray-Ban Display 出货——唯一规模商用的 EMG 接口。", CYAN),
    ]
    for i, (t, d, c) in enumerate(teams):
        left = Inches(0.45 + i * 3.2)
        box(s, left, Inches(1.55), Inches(3.05), Inches(3.9), WHITE, STROKE)
        rect(s, left, Inches(1.55), Inches(3.05), Inches(0.08), c)
        put(s, left + Inches(0.15), Inches(1.8), Inches(2.75), Inches(0.75), t, 15, c, True)
        put(s, left + Inches(0.15), Inches(2.6), Inches(2.75), Inches(2.7), d, 12, BODY)
    box(s, Inches(0.5), Inches(5.6), Inches(12.3), Inches(0.9), CARD, STROKE)
    put(s, Inches(0.7), Inches(5.75), Inches(11.9), Inches(0.6),
        "跟踪名单：CMU（MOMENT + TartanIMU）、Harvard + MIT 林肯实验室（UniTS）、Nokia Bell Labs（PRIMUS）、微软研究院（Babel）。",
        13, BODY)
    footer(s, 5, "TOP3 团队")

    # 6 TOP3 商用场景
    s = blank_slide(prs)
    header(s, "5  TOP3 商用场景", "平台能力已被产业验证的三个方向",
           "平台只提供接入 / 预测 / 分类接口；场景 Pack 由业务方交付。")
    scenes = [
        ("可穿戴运动健康", "传感器：IMU / EMG（手表、手环、腕带）。\n能力：手势 / 活动分类 + 生理指标预测。\n验证：Meta 腕带随 Ray-Ban Display 出货；LIMU-BERT-X 端侧 HAR。", RED),
        ("终端情境感知", "传感器：IMU（手机、耳机）。\n能力：骑行 / 驾驶 / 跌倒识别 + 状态预测。\n验证：LIMU-BERT-X 外卖配送全国部署，6 万骑手、1.1K 种机型。", ORANGE),
        ("工业与机器人", "传感器：IMU / TP（温度压力）。\n能力：工况分类 + 预测性维护 + 位姿估计。\n验证：TartanIMU 跨机器人平台，LoRA 1.1M 参数适配、200 FPS。", CYAN),
    ]
    for i, (t, d, c) in enumerate(scenes):
        left = Inches(0.45 + i * 3.2)
        box(s, left, Inches(1.55), Inches(3.05), Inches(3.9), WHITE, STROKE)
        rect(s, left, Inches(1.55), Inches(3.05), Inches(0.08), c)
        put(s, left + Inches(0.15), Inches(1.8), Inches(2.75), Inches(0.4), t, 16, c, True)
        put(s, left + Inches(0.15), Inches(2.35), Inches(2.75), Inches(2.9), d, 12, BODY)
    box(s, Inches(0.5), Inches(5.6), Inches(12.3), Inches(0.9), CARD, STROKE)
    put(s, Inches(0.7), Inches(5.75), Inches(11.9), Inches(0.6),
        "共性：三个场景都需要「预测 + 分类」同时在线，且设备 / 用户高度异构——正对应三大痛点。",
        14, RED, True)
    footer(s, 6, "TOP3 场景")

    # 7 痛点 → 方案
    s = blank_slide(prs)
    header(s, "6  痛点 → 方案", "每个痛点对应一类已验证的模型能力",
           "平台不做场景：只把模型能力收敛成 接入 → 特征 → 预测 / 分类。")
    rows = [
        ("多模态\n融合不足", "统一 patch token + 变元注意力：TimesFM-3 变元注意力、UniTS 任务 token；Babel 6 模态对齐作参照。\n接入规格统一时钟、量纲、质量位——对齐在进模型前完成。", RED),
        ("多设备\n自适应困难", "跨设备预训练先验：LIMU-BERT-X 覆盖 1.1K 机型、6 万人；Meta EMG 跨用户泛化。\nRevIN 实例归一 + LoRA 轻适配（TartanIMU 仅 1.1M 参数即适配新设备）。", ORANGE),
        ("多任务多场景\n泛化弱", "零样本基座：TimesFM-3 零样本预测；UniTS 零样本多任务（预测/分类/填补/异常）。\n统一预测 / 分类接口，新场景先零样本启动，再按需微调。", CYAN),
    ]
    for i, (k, v, c) in enumerate(rows):
        top = Inches(1.5 + i * 1.35)
        box(s, Inches(0.5), top, Inches(12.3), Inches(1.22), WHITE, STROKE)
        rect(s, Inches(0.5), top, Inches(0.1), Inches(1.22), c)
        put(s, Inches(0.78), top + Inches(0.14), Inches(2.8), Inches(0.95), k, 15, c, True)
        put(s, Inches(3.8), top + Inches(0.12), Inches(8.8), Inches(1.0), v, 12, BODY)
    box(s, Inches(0.5), Inches(5.65), Inches(12.3), Inches(0.85), CARD, STROKE)
    put(s, Inches(0.7), Inches(5.78), Inches(11.9), Inches(0.62),
        "平台边界：交付接入校验、共享特征、预测 / 分类双接口、评测工具链；不交付场景 Pack、整机、芯片。",
        13, RED, True)
    footer(s, 7, "痛点→方案")

    # 8 接口定义
    s = blank_slide(prs)
    header(s, "7  接口定义", "接入规格 + 预测接口 + 分类接口",
           "所有输入先过接入校验；预测与分类共用同一批 patch token。")
    box(s, Inches(0.5), Inches(1.5), Inches(12.3), Inches(2.0), WHITE, STROKE)
    rect(s, Inches(0.5), Inches(1.5), Inches(12.3), Inches(0.08), RED)
    put(s, Inches(0.7), Inches(1.75), Inches(11.9), Inches(0.35), "接入（数据规格） → 校验 → patch token", 16, RED, True)
    put(s, Inches(0.7), Inches(2.2), Inches(11.9), Inches(1.1),
        "数据规格 = {序列 id, 时间戳, 数值, 量纲单位, 质量位, 变元角色: 目标 / 仅历史 / 历史+未来}\n"
        "校验：时钟对齐、量纲归一、质量位过滤、变元上限 32、历史段 ≤ 15,360。不合格直接拒收。",
        13, BODY)
    box(s, Inches(0.5), Inches(3.7), Inches(6.1), Inches(2.4), WHITE, STROKE)
    rect(s, Inches(0.5), Inches(3.7), Inches(6.1), Inches(0.08), GREEN)
    put(s, Inches(0.7), Inches(3.95), Inches(5.7), Inches(0.35), "预测（历史段, 预测时长） → {分位数, 点预测}", 16, GREEN, True)
    put(s, Inches(0.7), Inches(4.45), Inches(5.7), Inches(1.4),
        "分位数：(V, H, 9) 十分位轨迹\n点预测：(V, H) 中位数 q50\n一次前向，CPM 非自回归；可拼接任意预测长度。",
        13, BODY)
    box(s, Inches(6.8), Inches(3.7), Inches(6.0), Inches(2.4), WHITE, STROKE)
    rect(s, Inches(6.8), Inches(3.7), Inches(6.0), Inches(0.08), CYAN)
    put(s, Inches(7.0), Inches(3.95), Inches(5.6), Inches(0.35), "分类（时间窗） → {类别, 置信度}", 16, CYAN, True)
    put(s, Inches(7.0), Inches(4.45), Inches(5.6), Inches(1.4),
        "类别：场景方注册的类别 id\n置信度：softmax 概率\n基于共享特征 + 任务 token / 线性探针。",
        13, BODY)
    footer(s, 8, "接口")

    # 9 代际演进
    s = blank_slide(prs)
    header(s, "8  代际演进", "Gen1 零样本基线 → Gen2 千人千面 → Gen3 Token 嵌入大模型",
           "接口不变，能力逐代升级；每代回答一个业务问题。")
    gens = [
        ("Gen1  零样本基线",
         "预训练底座直接上岗。\n预测零样本（TimesFM-3）；\n分类走规则 / 线性探针（UniTS / MOMENT）。\n回答：底座行不行 → 统一评测基线。", RED),
        ("Gen2  千人千面",
         "轻量适配到每设备、每用户。\nLoRA / 提示微调：TartanIMU 仅 1.1M 参数；\nMeta EMG 个性化后再 +16%。\n回答：换设备换人还行不行 → 画像进数据规格。", ORANGE),
        ("Gen3  Token 嵌入大模型",
         "物理 patch token 作为新模态注入大模型。\n预测 + 分类变成大模型的原生能力；\n平台从「提供模型」变成「提供 token 与接口」。\n回答：物理感知如何进入统一智能。", CYAN),
    ]
    for i, (t, d, c) in enumerate(gens):
        left = Inches(0.45 + i * 3.2)
        box(s, left, Inches(1.55), Inches(3.05), Inches(3.9), WHITE, STROKE)
        rect(s, left, Inches(1.55), Inches(3.05), Inches(0.08), c)
        put(s, left + Inches(0.15), Inches(1.8), Inches(2.75), Inches(0.4), t, 16, c, True)
        put(s, left + Inches(0.15), Inches(2.35), Inches(2.75), Inches(2.9), d, 12, BODY)
    box(s, Inches(0.5), Inches(5.6), Inches(12.3), Inches(0.9), CARD, STROKE)
    put(s, Inches(0.7), Inches(5.75), Inches(11.9), Inches(0.6),
        "演进原则：接口（接入 / 预测 / 分类）向下兼容；底座、权重、头可独立替换。",
        14, RED, True)
    footer(s, 9, "代际演进")

    try:
        prs.save(OUT)
        out = OUT
    except PermissionError:
        out = OUT.replace(".pptx", "-v2.pptx")
        prs.save(out)
    print("SAVED", out)
    print("SLIDES", len(prs.slides))


if __name__ == "__main__":
    build()


