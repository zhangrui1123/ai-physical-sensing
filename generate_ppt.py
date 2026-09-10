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
TOTAL = 8


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
        "华为内部立项  ·  平台工具化：接口明确 + 代际演进", 14, RED, True)
    put(s, Inches(0.7), Inches(1.7), Inches(12), Inches(0.7), "AI 物理感知平台", 40, NAVY, True)
    put(s, Inches(0.7), Inches(2.5), Inches(12), Inches(0.4),
        "Ingest 契约 → 共享表征 → predict() / classify()，不做场景应用", 18, BODY)
    box(s, Inches(0.7), Inches(3.1), Inches(11.9), Inches(1.2), CARD, STROKE)
    put(s, Inches(0.95), Inches(3.25), Inches(11.4), Inches(0.95),
        "平台只交付工具链：统一输入契约、两个推理函数、可替换底座。\n"
        "TimesFM-3 管预测，UniTS/MOMENT 管分类，LIMU-BERT-X 管 IMU 先验；底座可换，接口不变。",
        15, BODY)
    for i, (k, v) in enumerate([
        ("接口", "Ingest schema + predict + classify"),
        ("底座", "TimesFM-3 + UniTS/MOMENT + LIMU-BERT-X"),
        ("演进", "Gen1 零样本 → Gen2 指令微调"),
    ]):
        left = Inches(0.7 + i * 4.05)
        box(s, left, Inches(4.55), Inches(3.85), Inches(1.7), WHITE, STROKE)
        rect(s, left, Inches(4.55), Inches(3.85), Inches(0.08), RED)
        put(s, left + Inches(0.2), Inches(4.75), Inches(3.45), Inches(0.35), k, 13, RED, True)
        put(s, left + Inches(0.2), Inches(5.2), Inches(3.45), Inches(0.85), v, 15, NAVY)
    put(s, Inches(0.7), Inches(6.5), Inches(8.5), Inches(0.3), "来源：TimesFM / UniTS / MOMENT / LIMU-BERT-X 公开资料  ·  8 页平台版", 12, MUTED)
    put(s, Inches(10.3), Inches(6.5), Inches(2.5), Inches(0.3), f"1  /  {TOTAL}", 12, MUTED, align=PP_ALIGN.RIGHT)
    put(s, Inches(0.7), Inches(6.9), Inches(6), Inches(0.25), "Security Level: Internal", 10, MUTED)

    # 2 平台边界
    s = blank_slide(prs)
    header(s, "1  平台边界", "只做工具，不做场景；接口稳定，底座可换",
           "明确划分：平台负责 Ingest → 表征 → 推理；场景方负责业务标签与闭环。")
    rows = [
        ("平台交付", "PhysIngest 校验、共享表征、predict() / classify()、评测工具链"),
        ("平台不交付", "场景 Pack、整机 App、芯片、盘古、世界视频、把 TimesFM 当分类模型宣传"),
        ("底座可替换", "TimesFM-3 → 自研 / 更新版本；UniTS/MOMENT → 其他多任务模型"),
        ("接口不替换", "Ingest schema、predict() 返回结构、classify() 返回结构保持兼容"),
        ("输入", "target / past-only / past-future 变元；IMU / EMG / TP 只是变元"),
        ("输出", "预测：quantiles + point；分类：label + confidence；无场景语义"),
    ]
    for i, (k, v) in enumerate(rows):
        r, c = divmod(i, 2)
        left = Inches(0.5 + c * 6.35)
        top = Inches(1.5 + r * 1.5)
        box(s, left, top, Inches(6.15), Inches(1.35), WHITE, STROKE)
        put(s, left + Inches(0.22), top + Inches(0.18), Inches(5.7), Inches(0.35), k, 15, RED, True)
        put(s, left + Inches(0.22), top + Inches(0.58), Inches(5.7), Inches(0.6), v, 14, BODY)
    so_what(s, "平台不抢场景：场景方调 predict/classify，平台方保接口与底座演进。")
    footer(s, 2, "平台边界")

    # 3 接口定义
    s = blank_slide(prs)
    header(s, "2  接口定义", "Ingest 契约 + predict() + classify()",
           "所有输入先过 Ingest；预测与分类共用同一批 patch token。")
    box(s, Inches(0.5), Inches(1.5), Inches(12.3), Inches(2.0), WHITE, STROKE)
    rect(s, Inches(0.5), Inches(1.5), Inches(12.3), Inches(0.08), RED)
    put(s, Inches(0.7), Inches(1.75), Inches(11.9), Inches(0.35), "Ingest(schema) → 校验 → patch token", 16, RED, True)
    put(s, Inches(0.7), Inches(2.2), Inches(11.9), Inches(1.1),
        "schema = {series_id, timestamp, value, unit, quality_flag, variate_role: target/past-only/past-future}\n"
        "校验：时钟对齐、量纲归一、质量位过滤、变元上限 32、context ≤ 15,360。不合格直接拒收。",
        13, BODY)
    box(s, Inches(0.5), Inches(3.7), Inches(6.1), Inches(2.4), WHITE, STROKE)
    rect(s, Inches(0.5), Inches(3.7), Inches(6.1), Inches(0.08), GREEN)
    put(s, Inches(0.7), Inches(3.95), Inches(5.7), Inches(0.35), "predict(context, horizon) → {quantiles, point}", 16, GREEN, True)
    put(s, Inches(0.7), Inches(4.45), Inches(5.7), Inches(1.4),
        "quantiles: (V, H, 9) 十分位轨迹\npoint: (V, H) 中位数 q50\n一次前向，CPM 非自回归；可 stitching 任意 H。",
        13, BODY)
    box(s, Inches(6.8), Inches(3.7), Inches(6.0), Inches(2.4), WHITE, STROKE)
    rect(s, Inches(6.8), Inches(3.7), Inches(6.0), Inches(0.08), CYAN)
    put(s, Inches(7.0), Inches(3.95), Inches(5.6), Inches(0.35), "classify(window) → {label, confidence}", 16, CYAN, True)
    put(s, Inches(7.0), Inches(4.45), Inches(5.6), Inches(1.4),
        "label: 场景方注册的类别 id\nconfidence: softmax 概率\n基于共享表征 + 任务 token / linear probe。",
        13, BODY)
    footer(s, 3, "接口")

    # 4 预测实现
    s = blank_slide(prs)
    header(s, "3  预测实现", "TimesFM-3 式：CPM 一次前向出 9 分位数",
           "预测是平台原生能力；接口固定，底座可换。")
    preds = [
        ("输入", "context ≤ 15,360；变元 ≤ 32；target / past-only / past-future。"),
        ("特征", "pad32 → 去趋势 → RevIN → patch 192 → ResidualBlock → 1280 token。"),
        ("表征", "20 × MixingTransformer：因果时间注意力 + 变元注意力 + FFN。"),
        ("输出头", "Linear 1280→64×9；CPM mask horizon；逆 RevIN；stitching 任意 H。"),
        ("读出", "点预测 q50；概率区间 q10–q90；MAE + 覆盖率评测。"),
        ("替换点", "底座换 TimesFM-4 / 自研时，predict() 签名与返回结构不变。"),
    ]
    for i, (t, d) in enumerate(preds):
        r, c = divmod(i, 2)
        left = Inches(0.5 + c * 6.35)
        top = Inches(1.5 + r * 1.55)
        box(s, left, top, Inches(6.15), Inches(1.42), WHITE, STROKE)
        put(s, left + Inches(0.22), top + Inches(0.15), Inches(5.7), Inches(0.35), t, 16, RED, True)
        put(s, left + Inches(0.22), top + Inches(0.55), Inches(5.7), Inches(0.75), d, 13, BODY)
    footer(s, 4, "预测")

    # 5 分类实现
    s = blank_slide(prs)
    header(s, "4  分类实现", "共享表征 + 任务 token / linear probe",
           "分类不是 TimesFM 内建；平台用 UniTS / MOMENT 方式补。")
    box(s, Inches(0.5), Inches(1.5), Inches(6.1), Inches(4.6), WHITE, STROKE)
    rect(s, Inches(0.5), Inches(1.5), Inches(6.1), Inches(0.08), MUTED)
    put(s, Inches(0.7), Inches(1.75), Inches(5.7), Inches(0.4), "路 A  规则下游（零样本，示意）", 17, MUTED, True)
    put(s, Inches(0.7), Inches(2.3), Inches(5.7), Inches(3.5),
        "历史段：去趋势残差 z-score。\n|z|≥3 CRITICAL，≥2 WARNING，否则 NORMAL。\n\n"
        "未来段：观测是否落在 q10–q90 / q20–q80 外。\n\n"
        "不改权重，不能注册新类别，只能做异常示意。",
        14, BODY)
    box(s, Inches(6.8), Inches(1.5), Inches(6.0), Inches(4.6), WHITE, STROKE)
    rect(s, Inches(6.8), Inches(1.5), Inches(6.0), Inches(0.08), RED)
    put(s, Inches(7.0), Inches(1.75), Inches(5.6), Inches(0.4), "路 B  可学习分类（平台接口）", 17, RED, True)
    put(s, Inches(7.0), Inches(2.3), Inches(5.6), Inches(3.5),
        "UniTS：任务 token 把预测 / 分类 / 填补 / 异常收进同一套参数。\n"
        "MOMENT：掩码预训练编码器 + linear probe / 分类头。\n"
        "LIMU-BERT-X：IMU 掩码预训练，直接初始化 IMU 变元。\n\n"
        "场景方注册类别 → 平台微调 classify() → 返回 label + confidence。",
        14, BODY)
    footer(s, 5, "分类")

    # 6 代际演进
    s = blank_slide(prs)
    header(s, "5  代际演进", "Gen1 零样本 → Gen2 指令微调 → Gen3 端侧",
           "接口不变，底座与能力逐代升级。")
    gens = [
        ("Gen1  零样本基线", "TimesFM-3 + UniTS/MOMENT + LIMU-BERT-X\npredict 零样本；classify 规则 / LoRA\n评测：MAE + 覆盖率 + macro-F1", RED),
        ("Gen2  指令微调", "物理指令数据集：预测 / 分类 / 异常统一 prompt\n底座自研，摆脱 3.0 NC 权重\n评测：zero-shot → few-shot → full-shot", ORANGE),
        ("Gen3  端侧物理", "LIMU-BERT-X 式端侧初始化 + 量化\npredict/classify 在设备闭环\n评测：延迟 + 功耗 + 精度", CYAN),
    ]
    for i, (t, d, c) in enumerate(gens):
        left = Inches(0.45 + i * 3.2)
        box(s, left, Inches(1.55), Inches(3.05), Inches(3.9), WHITE, STROKE)
        rect(s, left, Inches(1.55), Inches(3.05), Inches(0.08), c)
        put(s, left + Inches(0.15), Inches(1.8), Inches(2.75), Inches(0.4), t, 16, c, True)
        put(s, left + Inches(0.15), Inches(2.35), Inches(2.75), Inches(2.9), d, 13, BODY)
    box(s, Inches(0.5), Inches(5.6), Inches(12.3), Inches(0.9), CARD, STROKE)
    put(s, Inches(0.7), Inches(5.75), Inches(11.9), Inches(0.6),
        "演进原则：接口（Ingest / predict / classify）向下兼容；底座、权重、头可独立替换。",
        14, RED, True)
    footer(s, 6, "代际演进")

    # 7 评测与许可
    s = blank_slide(prs)
    header(s, "6  评测与许可", "同一接口，同一评测，不同底座可对比",
           "平台提供统一评测工具链；许可边界随底座变化。")
    box(s, Inches(0.5), Inches(1.5), Inches(6.1), Inches(4.6), WHITE, STROKE)
    rect(s, Inches(0.5), Inches(1.5), Inches(6.1), Inches(0.08), GREEN)
    put(s, Inches(0.7), Inches(1.75), Inches(5.7), Inches(0.4), "评测接口", 17, GREEN, True)
    put(s, Inches(0.7), Inches(2.3), Inches(5.7), Inches(3.5),
        "predict：MAE、q10–q90 覆盖率、stitching 长 horizon 稳定性。\n"
        "classify：macro-F1、confusion matrix、注册类别一致性。\n"
        "底座对比：同一批数据，同一接口，换底座重跑。",
        14, BODY)
    box(s, Inches(6.8), Inches(1.5), Inches(6.0), Inches(4.6), WHITE, STROKE)
    rect(s, Inches(6.8), Inches(1.5), Inches(6.0), Inches(0.08), ORANGE)
    put(s, Inches(7.0), Inches(1.75), Inches(5.6), Inches(0.4), "许可边界", 17, ORANGE, True)
    put(s, Inches(7.0), Inches(2.3), Inches(5.6), Inches(3.5),
        "TimesFM-3 权重：NC，商用必须自训底座。\n"
        "TimesFM ≤ 2.5：Apache-2.0，可商用。\n"
        "UniTS / MOMENT / LIMU-BERT-X：公开可研究，商用前确认 license。\n\n"
        "平台不绑定任何 NC 权重；接口层与许可层解耦。",
        14, BODY)
    footer(s, 7, "评测与许可")

    # 8 决策
    s = blank_slide(prs)
    header(s, "决策", "平台交付接口与工具链，不交付场景语义",
           "接口固定，底座可换，代际演进向下兼容。")
    lines = [
        ("接口", "Ingest schema + predict() + classify()；版本化，向下兼容。"),
        ("预测", "TimesFM-3 式 CPM 一次前向；点预测 q50；评测 MAE + 覆盖率。"),
        ("分类", "UniTS 任务 token / MOMENT linear probe；场景方注册类别，平台微调。"),
        ("演进", "Gen1 零样本 → Gen2 指令微调 → Gen3 端侧；接口不变。"),
        ("许可", "3.0 权重 NC；2.5 及以前 Apache-2.0；平台层与许可层解耦。"),
        ("不做", "场景 Pack、整机、芯片、盘古、世界视频、把 TimesFM 当分类模型宣传。"),
    ]
    for i, (k, v) in enumerate(lines):
        top = Inches(1.5 + i * 0.8)
        box(s, Inches(0.5), top, Inches(12.3), Inches(0.72), WHITE if i % 2 == 0 else CARD, STROKE)
        put(s, Inches(0.7), top, Inches(1.6), Inches(0.72), k, 16, RED, True, valign=MSO_ANCHOR.MIDDLE)
        put(s, Inches(2.5), top, Inches(10.1), Inches(0.72), v, 14, NAVY, valign=MSO_ANCHOR.MIDDLE)
    footer(s, 8, "决策")

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


