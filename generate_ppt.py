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
        "华为内部立项  ·  融合 TimesFM-3 + UniTS/MOMENT + LIMU-BERT-X", 14, RED, True)
    put(s, Inches(0.7), Inches(1.7), Inches(12), Inches(0.7), "AI 物理感知平台", 40, NAVY, True)
    put(s, Inches(0.7), Inches(2.5), Inches(12), Inches(0.4),
        "一个物理时序底座，同时出预测与分类", 18, BODY)
    box(s, Inches(0.7), Inches(3.1), Inches(11.9), Inches(1.2), CARD, STROKE)
    put(s, Inches(0.95), Inches(3.25), Inches(11.4), Inches(0.95),
        "TimesFM-3 管零样本预测；UniTS / MOMENT 管多任务与分类；LIMU-BERT-X 管 IMU 表征。\n"
        "融合成 PhysFM：PhysIngest → 共享表征 → 预测头 + 分类头。不做场景应用、不做整机。",
        15, BODY)
    for i, (k, v) in enumerate([
        ("预测", "TimesFM-3 式 CPM · 9 分位数"),
        ("分类", "UniTS 任务 token / MOMENT 头"),
        ("IMU 先验", "LIMU-BERT-X · 143 万小时"),
    ]):
        left = Inches(0.7 + i * 4.05)
        box(s, left, Inches(4.55), Inches(3.85), Inches(1.7), WHITE, STROKE)
        rect(s, left, Inches(4.55), Inches(3.85), Inches(0.08), RED)
        put(s, left + Inches(0.2), Inches(4.75), Inches(3.45), Inches(0.35), k, 13, RED, True)
        put(s, left + Inches(0.2), Inches(5.2), Inches(3.45), Inches(0.85), v, 15, NAVY)
    put(s, Inches(0.7), Inches(6.5), Inches(8.5), Inches(0.3), "来源：TimesFM / UniTS / MOMENT / LIMU-BERT-X 公开资料  ·  8 页架构版", 12, MUTED)
    put(s, Inches(10.3), Inches(6.5), Inches(2.5), Inches(0.3), f"1  /  {TOTAL}", 12, MUTED, align=PP_ALIGN.RIGHT)
    put(s, Inches(0.7), Inches(6.9), Inches(6), Inches(0.25), "Security Level: Internal", 10, MUTED)

    # 2 规格
    s = blank_slide(prs)
    header(s, "1  规格", "融合底座：TimesFM 主干 + 双头 + IMU 预训练",
           "目标：同一批 patch token，同时给出分位数预测与类别 logits。")
    rows = [
        ("主干", "TimesFM-3：20 层 Mixing Transformer，d=1280，16 heads"),
        ("patch", "输入 32 / 输出 64；上下文上限 15,360；变元上限 32"),
        ("预测头", "Linear 1280→64×9，CPM 一次解码，median=q50"),
        ("分类头", "UniTS 任务 token 或 MOMENT linear-probe；可 LoRA 微调"),
        ("IMU 初始化", "LIMU-BERT-X：143 万小时、6 万人、1.1K 机型，端侧可用"),
        ("许可", "TimesFM-3 权重 NC；UniTS / MOMENT / LIMU-BERT 公开可研究。商用自训。"),
    ]
    for i, (k, v) in enumerate(rows):
        r, c = divmod(i, 2)
        left = Inches(0.5 + c * 6.35)
        top = Inches(1.5 + r * 1.5)
        box(s, left, top, Inches(6.15), Inches(1.35), WHITE, STROKE)
        put(s, left + Inches(0.22), top + Inches(0.18), Inches(5.7), Inches(0.35), k, 15, RED, True)
        put(s, left + Inches(0.22), top + Inches(0.58), Inches(5.7), Inches(0.6), v, 14, BODY)
    so_what(s, "TimesFM 原生没有分类头；要「预测+分类」必须融合 UniTS 式任务 token 或 MOMENT 式探针。")
    footer(s, 2, "规格")

    # 3 输入与流水线
    s = blank_slide(prs)
    header(s, "2  总览", "四段流水：输入 → 特征 → 表征 → 双头",
           "物理传感器只作为变元接入：target / past-only / past-future。")
    steps = [
        ("输入", "target (B,U,C)\npast-only (B,Vpo,C)\npast-future (B,W,C+H)", RED),
        ("特征提取", "pad32 · 去趋势 · 堆叠 V\n切 patch · RevIN · ResidualBlock", ORANGE),
        ("共享表征", "20 × MixingTransformer\n因果时间注意力 + 变元注意力", CYAN),
        ("双头", "预测：Linear 1280→64×9\n分类：任务 token / linear probe", GREEN),
    ]
    for i, (t, d, c) in enumerate(steps):
        left = Inches(0.45 + i * 3.2)
        box(s, left, Inches(1.55), Inches(3.05), Inches(3.15), WHITE, STROKE)
        rect(s, left, Inches(1.55), Inches(3.05), Inches(0.08), c)
        put(s, left + Inches(0.15), Inches(1.8), Inches(2.75), Inches(0.4), f"{i+1}. {t}", 16, c, True)
        put(s, left + Inches(0.15), Inches(2.35), Inches(2.75), Inches(2.1), d, 13, BODY)
    box(s, Inches(0.5), Inches(4.9), Inches(12.3), Inches(1.3), CARD, STROKE)
    put(s, Inches(0.7), Inches(5.05), Inches(11.9), Inches(0.3), "物理变元怎么接（不是应用场景）", 14, RED, True)
    put(s, Inches(0.7), Inches(5.4), Inches(11.9), Inches(0.7),
        "IMU / EMG 作 target 或 past-only；已知未来通道作 past-future，horizon 上保持可见。\n"
        "变元上限 32。无 schema / 时钟 / 质量位则拒收——这是 Ingest，TimesFM 不管。",
        13, BODY)
    footer(s, 3, "总览")

    # 4 特征提取
    s = blank_slide(prs)
    header(s, "3  特征提取", "数据是 patch token，不是逐步 RNN 状态",
           "每个时间 patch 拼成 192 维，ResidualBlock → 1280 维 token。")
    feats = [
        ("1 对齐", "context 左 pad 到 32 倍数。horizon 上目标全 mask；past-future 在 horizon 可见。"),
        ("2 变元堆叠", "target ⊕ past-only ⊕ past-future 沿 V 维拼接，供变元注意力交互。"),
        ("3 去趋势", "可选线性去趋势；去趋势后 std 明显更小（阈值 0.5）才减直线，预测后再加回。"),
        ("4 RevIN", "逐变元 running mean/std。horizon 用 CPM：mask 目标 patch，用上下文统计并可迭代修正。"),
        ("5 Patch 嵌入", "当前 32 + roll 未来 64 + mask 96 → concat 192 → ResidualBlock（两层线性+ReLU+残差）→ d=1280。"),
    ]
    for i, (t, d) in enumerate(feats):
        top = Inches(1.48 + i * 0.9)
        box(s, Inches(0.5), top, Inches(12.3), Inches(0.82), WHITE if i % 2 == 0 else CARD, STROKE)
        put(s, Inches(0.7), top, Inches(2.4), Inches(0.82), t, 16, RED, True, valign=MSO_ANCHOR.MIDDLE)
        put(s, Inches(3.2), top, Inches(9.4), Inches(0.82), d, 14, BODY, valign=MSO_ANCHOR.MIDDLE)
    footer(s, 4, "特征提取")

    # 5 Mixing Transformer
    s = blank_slide(prs)
    header(s, "4  表征", "Mixing Transformer 学的是条件表征，预测与分类共用",
           "每层对 (B, V, N, 1280) 做：时间注意力 → 变元注意力 → FFN。堆叠 20 层。")
    cards = [
        (CYAN, "时间注意力",
         "因果 + RoPE。\n每个变元独立：第 t 个 patch 只能看 ≤ t。\n建模趋势、季节、水平漂移。"),
        (ORANGE, "变元注意力",
         "同一时刻、跨通道、非因果。\n目标与协变量在同一 patch 交换信息。\n这是 3.0 相对 2.x 单变量的核心。"),
        (GREEN, "FFN",
         "RMSNorm + ReLU MLP + 残差。\n输出条件表征，同时供预测头与分类头。\nUniTS / MOMENT 证明这套表征可做分类。"),
    ]
    for i, (c, t, d) in enumerate(cards):
        left = Inches(0.5 + i * 4.2)
        box(s, left, Inches(1.55), Inches(4.0), Inches(3.55), WHITE, STROKE)
        rect(s, left, Inches(1.55), Inches(0.1), Inches(3.55), c)
        put(s, left + Inches(0.25), Inches(1.75), Inches(3.5), Inches(0.45), t, 18, c, True)
        put(s, left + Inches(0.25), Inches(2.4), Inches(3.5), Inches(2.4), d, 15, BODY)
    so_what(s, "表征层不动 TimesFM；分类能力来自在表征上接 UniTS 任务 token 或 MOMENT 探针。")
    footer(s, 5, "表征")

    # 6 预测
    s = blank_slide(prs)
    header(s, "5  预测", "decode() 非自回归：整段 context+horizon 只走一遍",
           "TimesFM3Torch.decode() → forward() → TimesFM3Forecaster.predict_batch()。")
    preds = [
        ("输出头", "Linear(1280 → 64×9)。每个输入 patch 预测未来 64 步、9 个分位数。"),
        ("CPM", "horizon 目标 patch 先 mask。用上下文统计量做 RevIN，并可迭代用模型估计修正。"),
        ("逆变换", "分位数在 RevIN 空间；逆归一化后再加回去趋势。"),
        ("Stitching", "相邻窗口重叠 32 点，拼出任意 horizon（可超过 64）。"),
        ("读出", "点预测：quantiles[..., 4]（0.5）。概率：9 条分位轨迹，做区间与校准。"),
        ("Eval", "点预测 MAE；q10–q90 覆盖率。对照 TimesFM-3 单变量 / 多变量两种模式。"),
    ]
    for i, (t, d) in enumerate(preds):
        r, c = divmod(i, 2)
        left = Inches(0.5 + c * 6.35)
        top = Inches(1.5 + r * 1.55)
        box(s, left, top, Inches(6.15), Inches(1.42), WHITE, STROKE)
        put(s, left + Inches(0.22), top + Inches(0.15), Inches(5.7), Inches(0.35), t, 16, RED, True)
        put(s, left + Inches(0.22), top + Inches(0.55), Inches(5.7), Inches(0.75), d, 13, BODY)
    footer(s, 6, "预测")

    # 7 分类
    s = blank_slide(prs)
    header(s, "6  分类", "分类不是 TimesFM 内建；融合 UniTS / MOMENT / LIMU-BERT-X",
           "规则只做异常示意；可学习分类必须接在共享表征上。")
    box(s, Inches(0.5), Inches(1.5), Inches(6.1), Inches(4.6), WHITE, STROKE)
    rect(s, Inches(0.5), Inches(1.5), Inches(6.1), Inches(0.08), MUTED)
    put(s, Inches(0.7), Inches(1.75), Inches(5.7), Inches(0.4), "路 A  规则下游（零样本）", 17, MUTED, True)
    put(s, Inches(0.7), Inches(2.3), Inches(5.7), Inches(3.5),
        "历史段：去趋势残差 z-score。\n|z|≥3 CRITICAL，≥2 WARNING，否则 NORMAL。\n\n"
        "未来段：观测是否落在 q10–q90 / q20–q80 外。\n区间外 → 异常。\n\n"
        "不改权重，不能当「可学习分类」。",
        14, BODY)
    box(s, Inches(6.8), Inches(1.5), Inches(6.0), Inches(4.6), WHITE, STROKE)
    rect(s, Inches(6.8), Inches(1.5), Inches(6.0), Inches(0.08), RED)
    put(s, Inches(7.0), Inches(1.75), Inches(5.6), Inches(0.4), "路 B  共享表征 + 分类头", 17, RED, True)
    put(s, Inches(7.0), Inches(2.3), Inches(5.6), Inches(3.5),
        "UniTS：任务 token 把预测 / 分类 / 填补 / 异常收进同一套参数。\n"
        "MOMENT：掩码预训练编码器 + linear probe / 分类头。\n"
        "LIMU-BERT-X：IMU 掩码预训练，直接初始化 IMU 变元。\n\n"
        "物理信号分类走这条；LoRA 微调参照 timesfm-forecasting/examples/finetuning/。",
        14, BODY)
    footer(s, 7, "分类")

    # 8 决策
    s = blank_slide(prs)
    header(s, "决策", "平台交付「预测+分类」融合底座，不交付应用 Pack",
           "TimesFM 管预测，UniTS / MOMENT 管分类，LIMU-BERT-X 管 IMU。")
    lines = [
        ("主干", "PhysIngest → pad/RevIN/patch token → 20 层 Mixing Transformer → 双头。"),
        ("预测", "TimesFM-3 式 CPM 一次前向；点预测 q50；评测 MAE + 分位覆盖率。"),
        ("分类", "UniTS 任务 token / MOMENT linear probe；可 LoRA。对照 LIMU-BERT-X。"),
        ("IMU", "用 LIMU-BERT-X 初始化 IMU 变元；TartanIMU 作运动估计参照。"),
        ("许可", "TimesFM-3 权重 NC；UniTS / MOMENT / LIMU-BERT 公开。商用自训。"),
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


