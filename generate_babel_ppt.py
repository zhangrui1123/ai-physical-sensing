# -*- coding: utf-8 -*-
"""Babel 论文解读 PPT（SenSys 2025）。"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree

W, H = Inches(13.333), Inches(7.5)
BG = RGBColor(0x0B, 0x12, 0x20)
CARD = RGBColor(0x13, 0x1C, 0x2E)
CARD2 = RGBColor(0x18, 0x24, 0x3A)
STROKE = RGBColor(0x2A, 0x3A, 0x55)
CYAN = RGBColor(0x38, 0xBD, 0xF8)
INDIGO = RGBColor(0x81, 0x8C, 0xF8)
AMBER = RGBColor(0xFB, 0xBF, 0x24)
GREEN = RGBColor(0x34, 0xD3, 0x99)
ROSE = RGBColor(0xFB, 0x71, 0x85)
WHITE = RGBColor(0xF8, 0xFA, 0xFC)
MUTED = RGBColor(0x94, 0xA3, 0xB8)
SOFT = RGBColor(0xCB, 0xD5, 0xE1)
FONT = "Microsoft YaHei"
TOTAL = 18


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
        sh.line.width = Pt(1)
    try:
        sh.adjustments[0] = 0.08
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


def bg(slide):
    rect(slide, 0, 0, W, H, BG)


def footer(slide, page, section=""):
    rect(slide, 0, Inches(7.28), W, Inches(0.22), RGBColor(0x08, 0x0D, 0x16))
    put(
        slide, Inches(0.4), Inches(7.28), Inches(9.4), Inches(0.22),
        "Babel  ·  SenSys 2025  ·  " + section if section else "Babel  ·  SenSys 2025",
        9, MUTED, valign=MSO_ANCHOR.MIDDLE,
    )
    put(
        slide, Inches(11.3), Inches(7.28), Inches(1.6), Inches(0.22),
        f"{page}  /  {TOTAL}", 9, MUTED, align=PP_ALIGN.RIGHT, valign=MSO_ANCHOR.MIDDLE,
    )


def header(slide, kicker, title, subtitle=None):
    rect(slide, 0, 0, Inches(0.12), H, CYAN)
    put(slide, Inches(0.45), Inches(0.14), Inches(12.4), Inches(0.26), kicker, 11, CYAN, True)
    put(slide, Inches(0.45), Inches(0.38), Inches(12.5), Inches(0.52), title, 21, WHITE, True)
    if subtitle:
        put(slide, Inches(0.45), Inches(0.88), Inches(12.5), Inches(0.32), subtitle, 13, MUTED)


def so_what(slide, text):
    box(slide, Inches(0.4), Inches(6.48), Inches(12.5), Inches(0.72), CARD2, AMBER)
    rect(slide, Inches(0.4), Inches(6.48), Inches(0.1), Inches(0.72), AMBER)
    put(slide, Inches(0.7), Inches(6.48), Inches(12.0), Inches(0.72),
        "所以  ·  " + text, 13, WHITE, True, valign=MSO_ANCHOR.MIDDLE)


def build():
    prs = Presentation()
    prs.slide_width, prs.slide_height = W, H
    blank = prs.slide_layouts[6]

    # 1 Cover
    s = prs.slides.add_slide(blank)
    bg(s)
    rect(s, 0, 0, Inches(0.16), H, CYAN)
    put(s, Inches(0.7), Inches(1.15), Inches(12), Inches(0.3),
        "论文解读  ·  ACM SenSys 2025  ·  arXiv:2407.17777", 13, AMBER, True)
    put(s, Inches(0.7), Inches(1.6), Inches(12), Inches(0.7), "Babel", 40, WHITE, True)
    put(s, Inches(0.7), Inches(2.35), Inches(12.2), Inches(0.7),
        "A Scalable Pre-trained Model for Multi-Modal Sensing\nvia Expandable Modality Alignment",
        18, SOFT)
    put(s, Inches(0.7), Inches(3.25), Inches(12), Inches(0.55),
        "Dai, Jiang, Yang, Cao, Li, Banerjee, Qiu  ·  UW–Madison / MSR / HKUST", 14, MUTED)
    box(s, Inches(0.7), Inches(4.05), Inches(11.9), Inches(1.35), CARD, STROKE)
    put(s, Inches(0.95), Inches(4.22), Inches(11.4), Inches(1.05),
        "核心命题：传感领域没有 CLIP 那种「全模态成对数据」。\n"
        "Babel 把 N 路对齐拆成一串二元对齐，用部分成对数据对齐 6 种传感模态。",
        15, SOFT)
    for i, (k, v) in enumerate([
        ("想法", "共享模态当桥，不必 N 元组"),
        ("方法", "模态塔 + 原型网 + 自适应权重"),
        ("结果", "单模态 +12%  ·  融合最高 +22%"),
    ]):
        left = Inches(0.7 + i * 4.05)
        box(s, left, Inches(5.6), Inches(3.85), Inches(1.15), CARD2, STROKE)
        put(s, left + Inches(0.2), Inches(5.7), Inches(3.45), Inches(0.32), k, 12, CYAN, True)
        put(s, left + Inches(0.2), Inches(6.05), Inches(3.45), Inches(0.5), v, 13, WHITE)
    put(s, Inches(10.3), Inches(6.95), Inches(2.5), Inches(0.3),
        f"1  /  {TOTAL}", 12, MUTED, align=PP_ALIGN.RIGHT)

    # 2 One-pager
    s = prs.slides.add_slide(blank)
    bg(s)
    header(s, "0  一页读懂", "Babel 要解决的不是「再做一个更大的融合网络」",
           "而是：成对数据极少时，如何把多种物理传感器拉进同一个表示空间。")
    items = [
        ("问题", "CLIP 用 4 亿图文对；公开传感成对数据只有 600–4.2 万。没有 6 模态同时成对的数据集。"),
        ("观察", "① 单模态编码器已经很强（如 LIMU-BERT）。② 不同数据集共享桥接模态（骨架、视频、Wi‑Fi）。"),
        ("做法", "N 路对齐 → 一串二元对齐。冻结编码器，只训对齐头；用原型网络避免灾难遗忘。"),
        ("交付", "预训练后，任选 1 路或多路模态做下游，不必为每种组合重训主干。"),
    ]
    for i, (k, v) in enumerate(items):
        top = Inches(1.35 + i * 1.2)
        box(s, Inches(0.4), top, Inches(12.5), Inches(1.08), CARD, STROKE)
        put(s, Inches(0.65), top + Inches(0.12), Inches(1.6), Inches(0.84), k, 16, CYAN, True, valign=MSO_ANCHOR.MIDDLE)
        put(s, Inches(2.4), top + Inches(0.18), Inches(10.2), Inches(0.75), v, 14, SOFT, valign=MSO_ANCHOR.MIDDLE)
    footer(s, 2, "总览")

    # 3 Problem
    s = prs.slides.add_slide(blank)
    bg(s)
    header(s, "1  想法  ·  问题", "把 CLIP 那套直接搬到传感，会卡在两层数据稀缺上",
           "现有工作要么只对齐 2–3 路，要么把 IMU 塞进 VLM 后几乎不可用。")
    box(s, Inches(0.4), Inches(1.35), Inches(6.2), Inches(4.95), CARD, ROSE)
    put(s, Inches(0.65), Inches(1.5), Inches(5.7), Inches(0.4), "稀缺 A  ·  成对样本不够", 16, ROSE, True)
    put(s, Inches(0.65), Inches(2.1), Inches(5.7), Inches(3.9),
        "CLIP 预训练约 4×10⁸ 图文对。\n\n"
        "公开多模态传感数据集通常只有 600–42,000 对。\n\n"
        "Wi‑Fi CSI、mmWave、LiDAR 采集要专用硬件，不可能靠互联网爬取。\n\n"
        "从零训对齐编码器，样本不够收敛。",
        14, SOFT)
    box(s, Inches(6.8), Inches(1.35), Inches(6.1), Inches(4.95), CARD, ROSE)
    put(s, Inches(7.05), Inches(1.5), Inches(5.65), Inches(0.4), "稀缺 B  ·  多模态同时成对不存在", 16, ROSE, True)
    put(s, Inches(7.05), Inches(2.1), Inches(5.65), Inches(3.9),
        "对齐 N 路通常需要 N 元组样本。\n\n"
        "公开数据最多覆盖子集：没有「六模态同时出现」的包。\n\n"
        "Cosmo 只能在同一数据集内融 RGB/深度/IMU。\n"
        "ImageBind / OneLLM 对 IMU 几乎只吃 Ego4D，HAR 上约 5–7%。",
        14, SOFT)
    footer(s, 3, "想法")

    # 4 Two observations
    s = prs.slides.add_slide(blank)
    bg(s)
    header(s, "1  想法  ·  关键观察", "数据不够，但「单模态编码器」和「共享桥」够用",
           "这两条观察直接决定了后面三块算法。")
    box(s, Inches(0.4), Inches(1.35), Inches(6.2), Inches(3.85), CARD, CYAN)
    put(s, Inches(0.65), Inches(1.52), Inches(5.7), Inches(0.45), "观察 1", 14, CYAN, True)
    put(s, Inches(0.65), Inches(2.05), Inches(5.7), Inches(1.0),
        "传感信号有明确物理含义，成熟单模态编码器已经抽出可迁移特征。", 16, WHITE, True)
    put(s, Inches(0.65), Inches(3.2), Inches(5.7), Inches(1.7),
        "因此不必从随机初始化学编码器。\n冻结预训练塔，只训很薄的对齐 MLP（PEFT / LiT 思路）。\n成对样本需求大幅下降。",
        14, SOFT)
    box(s, Inches(6.8), Inches(1.35), Inches(6.1), Inches(3.85), CARD, AMBER)
    put(s, Inches(7.05), Inches(1.52), Inches(5.65), Inches(0.45), "观察 2", 14, AMBER, True)
    put(s, Inches(7.05), Inches(2.05), Inches(5.65), Inches(1.0),
        "几乎没有三路以上全配对，但大量二元数据集共享同一模态。", 16, WHITE, True)
    put(s, Inches(7.05), Inches(3.2), Inches(5.65), Inches(1.7),
        "共享模态可以当桥：IMU–骨架、骨架–视频、骨架–Wi‑Fi、Wi‑Fi–mmWave、视频–LiDAR。\n五次二元对齐，覆盖六路模态。",
        14, SOFT)
    so_what(s, "把 N 模态对齐改写成「沿桥生长的二元对齐序列」。这就是 expandable modality alignment。")
    footer(s, 4, "想法")

    # 5 Bridge datasets
    s = prs.slides.add_slide(blank)
    bg(s)
    header(s, "1  想法  ·  数据桥", "六模态不靠一个超级数据集，靠五座「两两配对」的桥",
           "预训练自监督，不用活动标签。深度被转成骨架，文中 depth ≈ skeleton。")
    chain = [
        ("IMU", "骨架", "UTD-MHAD", "613"),
        ("骨架", "视频", "Kinetics-400", "23.5 万"),
        ("骨架", "Wi-Fi", "OPERANet", "2.5 万"),
        ("Wi-Fi", "mmWave", "XRF55", "3.0 万"),
        ("视频", "LiDAR", "MM-Fi", "1.8 万"),
    ]
    for i, (a, b, ds, n) in enumerate(chain):
        left = Inches(0.35 + i * 2.58)
        box(s, left, Inches(1.4), Inches(2.45), Inches(3.15), CARD, STROKE)
        put(s, left + Inches(0.1), Inches(1.55), Inches(2.25), Inches(0.7), a + "\n↔  " + b, 15, WHITE, True, PP_ALIGN.CENTER)
        put(s, left + Inches(0.1), Inches(2.4), Inches(2.25), Inches(0.9), ds, 13, CYAN, True, PP_ALIGN.CENTER)
        put(s, left + Inches(0.1), Inches(3.35), Inches(2.25), Inches(0.9), n + " 训练对", 12, MUTED, align=PP_ALIGN.CENTER)
    box(s, Inches(0.4), Inches(4.75), Inches(12.5), Inches(1.55), CARD2, STROKE)
    put(s, Inches(0.65), Inches(4.9), Inches(12.1), Inches(1.25),
        "论文默认生长顺序（§7.4）：IMU+骨架 → 视频挂到骨架 → Wi‑Fi 挂到骨架 → mmWave 挂到 Wi‑Fi → LiDAR 挂到视频。\n"
        "UTD-MHAD 另做 600× 增强（降采样模拟不同采样率 + 动作截断模拟不完整活动）。",
        14, SOFT)
    footer(s, 5, "想法")

    # 6 Architecture overview
    s = prs.slides.add_slide(blank)
    bg(s)
    header(s, "2  架构  ·  总图", "三块积木：模态塔、可扩展主干、自适应训练",
           "对应论文 §4 / §5 / §6。部署时任选已对齐模态，接一个很浅的下游头。")
    blocks = [
        ("① 预训练模态塔", "冻结的单模态编码器\n+ 可训的 Concept Alignment MLP",
         "回答：成对样本太少怎么办？\n不重训编码器。"),
        ("② 可扩展网络", "Trunk + Branch + Junction\n共享 Prototype Network",
         "回答：没有 N 元组怎么办？\n二元生长，桥接遗忘。"),
        ("③ 自适应训练", "对比损失按梯度范数加权\n原型网用 EMA + 蒸馏",
         "回答：新模态会冲掉旧对齐吗？\n弱模态靠向强模态。"),
    ]
    for i, (t, a, b) in enumerate(blocks):
        left = Inches(0.4 + i * 4.25)
        box(s, left, Inches(1.35), Inches(4.05), Inches(4.95), CARD, STROKE)
        put(s, left + Inches(0.2), Inches(1.55), Inches(3.65), Inches(0.7), t, 16, CYAN, True)
        put(s, left + Inches(0.2), Inches(2.4), Inches(3.65), Inches(1.5), a, 14, WHITE)
        put(s, left + Inches(0.2), Inches(4.15), Inches(3.65), Inches(1.8), b, 14, AMBER)
    footer(s, 6, "架构")

    # 7 Modality tower
    s = prs.slides.add_slide(blank)
    bg(s)
    header(s, "2  架构  ·  模态塔", "每路传感器 = 领域编码器（冻）+ 对齐 MLP（训）",
           "选型原则：有物理含义的用信号处理；噪声大的用深度学习；避免域差过大的预训练权重。")
    encs = [
        ("IMU", "LIMU-BERT", "多 IMU 数据集预训练", "72"),
        ("骨架", "ST-GCN", "NTU-RGBD", "60"),
        ("视频", "ResNet3D", "Kinetics-400", "512"),
        ("Wi-Fi", "ViT + CNN-GRU", "塔增强，UT-HAR", "900"),
        ("mmWave", "多普勒/角度 FFT + ResNet18", "信号处理 + 空间网", "1024"),
        ("LiDAR", "Point Transformer + ST-GCN", "ModelNet40 + NTU", "60"),
    ]
    for i, (m, enc, pre, dim) in enumerate(encs):
        r, c = divmod(i, 3)
        left = Inches(0.4 + c * 4.25)
        top = Inches(1.32 + r * 2.4)
        box(s, left, top, Inches(4.05), Inches(2.2), CARD, STROKE)
        put(s, left + Inches(0.2), top + Inches(0.12), Inches(3.65), Inches(0.35), m, 16, CYAN, True)
        put(s, left + Inches(0.2), top + Inches(0.52), Inches(3.65), Inches(0.55), enc, 13, WHITE, True)
        put(s, left + Inches(0.2), top + Inches(1.1), Inches(3.65), Inches(0.5), pre, 12, MUTED)
        put(s, left + Inches(0.2), top + Inches(1.6), Inches(3.65), Inches(0.4), "对齐头输入维  " + dim, 12, AMBER)
    footer(s, 7, "架构")

    # 8 Contrastive
    s = prs.slides.add_slide(blank)
    bg(s)
    header(s, "2  架构  ·  二元对齐", "对齐时只更新 Concept Alignment，对比损失是双向 InfoNCE",
           "batch=256 → 256 个正对、65,280 个负对。温度 τ = 0.07。")
    box(s, Inches(0.4), Inches(1.35), Inches(7.9), Inches(4.95), CARD, STROKE)
    put(s, Inches(0.65), Inches(1.5), Inches(7.4), Inches(0.4), "一次迭代在做什么", 15, CYAN, True)
    put(s, Inches(0.65), Inches(2.05), Inches(7.4), Inches(3.95),
        "1. 从成对数据集抽 batch：同一活动、同时刻的 (χα, χβ)。\n\n"
        "2. 正对：同一序号的两路信号。负对：batch 内交叉配对。\n\n"
        "3. 冻编码器，两路过各自对齐 MLP，算余弦相似度。\n\n"
        "4. 双向损失 Lαβ = (Lα←β + Lβ←α) / 2，反传只进对齐头。\n\n"
        "下游：各模态对齐 embedding 直接拼接，接 2 层 MLP，one-shot 即可。",
        14, SOFT)
    box(s, Inches(8.5), Inches(1.35), Inches(4.4), Inches(4.95), CARD2, STROKE)
    put(s, Inches(8.7), Inches(1.5), Inches(4.0), Inches(0.4), "Wi-Fi 为何要塔增强", 14, AMBER, True)
    put(s, Inches(8.7), Inches(2.1), Inches(4.0), Inches(3.9),
        "CSI 噪声大，单一预训练编码器不够。\n\n"
        "同一模态建两个塔（ViT 与 CNN-GRU），用同一样本的正对把两个编码器对齐，相当于弱学习器集成。\n\n"
        "这是针对「脏传感」的工程补丁，不是装饰。",
        13, SOFT)
    footer(s, 8, "架构")

    # 9 Expandable + prototype
    s = prs.slides.add_slide(blank)
    bg(s)
    header(s, "2  架构  ·  生长", "新模态不是重训全集，而是挂到主干的「路口」上",
           "直接更新路口塔会冲掉已经对齐的旁路。所以要有共享的原型网络。")
    steps = [
        ("Trunk", "先用 Eαβ 对齐 α 与 β，得到主干 Hαβ。"),
        ("Junction", "新模态 κ 只与 α 成对 → α 是路口。取出已训的 Γα。"),
        ("Branch", "新建 Γκ，用 Eακ 把分支对齐到路口。称为 growth。"),
        ("Prototype Υ", "接在每个对齐头之后、全模态共享的 2–4 层 MLP。协调旧知识与新洞察。"),
    ]
    for i, (k, v) in enumerate(steps):
        top = Inches(1.32 + i * 1.18)
        box(s, Inches(0.4), top, Inches(12.5), Inches(1.05), CARD, STROKE)
        put(s, Inches(0.65), top, Inches(2.6), Inches(1.05), k, 16, CYAN, True, valign=MSO_ANCHOR.MIDDLE)
        put(s, Inches(3.4), top, Inches(9.2), Inches(1.05), v, 15, SOFT, valign=MSO_ANCHOR.MIDDLE)
    footer(s, 9, "架构")

    # 10 Adaptive training
    s = prs.slides.add_slide(blank)
    bg(s)
    header(s, "2  架构  ·  自适应", "生长时谁该被谁拉近，由梯度范数当场决定",
           "不可靠的新模态应靠向已对齐主干，而不是把主干拽歪。")
    box(s, Inches(0.4), Inches(1.32), Inches(7.9), Inches(4.98), CARD, STROKE)
    put(s, Inches(0.65), Inches(1.48), Inches(7.4), Inches(0.4), "加权双向对比  （公式 5–7）", 15, CYAN, True)
    put(s, Inches(0.65), Inches(2.05), Inches(7.4), Inches(3.95),
        "L = [ wα←β · Lα←β + wβ←α · Lβ←α ] / 2\n\n"
        "权重 w ∝ 1 / ‖∇‖，再归一化使两者之和为 1。\n\n"
        "直觉：某方向梯度小 → 该侧更稳，给更大权重，让另一侧靠过来。\n"
        "分支梯度大（离对齐还远）→ 加大路口权重，把新模态「拉进」主干。\n\n"
        "论文 Fig.6：Wi‑Fi 并入骨架时，骨架权重起初接近 1，约 6000 step 后才开始双向交换知识。",
        14, SOFT)
    box(s, Inches(8.5), Inches(1.32), Inches(4.4), Inches(4.98), CARD2, STROKE)
    put(s, Inches(8.7), Inches(1.48), Inches(4.0), Inches(0.4), "原型网怎么更新", 14, AMBER, True)
    put(s, Inches(8.7), Inches(2.1), Inches(4.0), Inches(3.9),
        "指数滑动平均（EMA）：慢速吸收新信息，保住已有原型。\n\n"
        "辅以知识蒸馏，减少生长时的灾难遗忘。\n\n"
        "结构刻意保持简单，才能跨数据集、跨任务反复增强同一套共享参数。",
        13, SOFT)
    footer(s, 10, "架构")

    # 11 Innovations
    s = prs.slides.add_slide(blank)
    bg(s)
    header(s, "3  创新点", "相对 Cosmo / ImageBind / CLIP，Babel 真正新的是「可扩展」而不是「对比学习」",
           "对比学习本身不是贡献。贡献是：在传感数据约束下让对比学习能长到六路。")
    innos = [
        ("1", "可扩展对齐范式",
         "首次把传感基础模型的 N 路对齐写成二元生长链，部署时可任选子集，不必为组合重训。"),
        ("2", "预训练模态塔",
         "复用 LIMU-BERT 等单模态塔 + 冻结编码器。没有塔，论文写明：样本太少，训练不收敛。"),
        ("3", "原型网络抗遗忘",
         "共享 MLP 编码已对齐的公共特征。去掉它，旧模态相对掉点约 44.7%（UTD-MHAD）。"),
        ("4", "梯度自适应权重",
         "按 batch 梯度范数分配谁靠谁。去掉自适应，整体最多掉 7.2%。生长顺序变得不敏感。"),
        ("5", "脏模态塔增强",
         "Wi‑Fi 用双编码器集成，避免单一弱塔拖垮桥。"),
        ("6", "接生成与 LLM",
         "对齐空间可接到 unCLIP（传感成像）和 Video-LLaMA（IMU 不重训即可被语言理解）。"),
    ]
    for i, (n, t, d) in enumerate(innos):
        r, c = divmod(i, 3)
        left = Inches(0.4 + c * 4.25)
        top = Inches(1.32 + r * 2.7)
        box(s, left, top, Inches(4.05), Inches(2.5), CARD, STROKE)
        put(s, left + Inches(0.18), top + Inches(0.12), Inches(3.7), Inches(0.3), n, 12, AMBER, True)
        put(s, left + Inches(0.18), top + Inches(0.45), Inches(3.7), Inches(0.5), t, 15, WHITE, True)
        put(s, left + Inches(0.18), top + Inches(1.05), Inches(3.7), Inches(1.25), d, 12, SOFT)
    footer(s, 11, "创新")

    # 12 Setup
    s = prs.slides.add_slide(blank)
    bg(s)
    header(s, "4  实验  ·  怎么评才公平", "One-shot HAR：每类只用 1 个样本微调分类头，测的是表征而不是刷标签",
           "传感标注贵，one-shot 比全监督更能说明「这是不是基础模型」。")
    box(s, Inches(0.4), Inches(1.32), Inches(6.2), Inches(5.0), CARD, STROKE)
    put(s, Inches(0.65), Inches(1.48), Inches(5.7), Inches(0.4), "数据", 15, CYAN, True)
    put(s, Inches(0.65), Inches(2.0), Inches(5.7), Inches(4.0),
        "预训练 5 集：UTD-MHAD、MM-Fi、OPERANet、XRF55、Kinetics-400。\n\n"
        "评测 8 集：\n"
        "域内 4 = 上述除 Kinetics 外的测试对。\n"
        "域外 4 = UCI、Widar3.0、mRI、MSRAction3D（预训练完全没见过）。\n\n"
        "训练：AdamW，lr=1e-4，bs=256，两张 A100 约 20 小时对齐六模态。",
        14, SOFT)
    box(s, Inches(6.8), Inches(1.32), Inches(6.1), Inches(5.0), CARD, STROKE)
    put(s, Inches(7.05), Inches(1.48), Inches(5.65), Inches(0.4), "基线", 15, AMBER, True)
    put(s, Inches(7.05), Inches(2.0), Inches(5.65), Inches(4.0),
        "单模态 SOTA：\nLIMU-BERT / SenseFi / MARS / PointTransformer / ResNet3D / ST-GCN\n\n"
        "多模态融合：Cosmo（需同一数据集内全模态共存）\n\n"
        "MLLM：OneLLM（Meta-Transformer）、M4（ImageBind）\n"
        "它们实际几乎只支持 IMU，且跨域很弱。",
        14, SOFT)
    footer(s, 12, "实验")

    # 13 Single-modal numbers
    s = prs.slides.add_slide(blank)
    bg(s)
    header(s, "4  结果  ·  单模态", "对齐之后，即使只用一路传感器，弱模态也能从强模态「借」到表示",
           "域内六模态平均 +12%，最高约 +20%。视频已经很强，只再涨约 2%。")
    rows = [
        ("IMU  ·  UTD-MHAD", "20.19% → 31.77%", "vs LIMU-BERT"),
        ("Wi-Fi", "+10.74%", "相对 SenseFi 量级"),
        ("mmWave  ·  XRF55", "30.32% → 50.30%", "弱射频收益最大"),
        ("LiDAR  ·  MM-Fi", "28.43% → 43.91%", "点云同样受益"),
        ("Video", "约 +2%", "强模态可借出、少借入"),
        ("域外 IMU", "74.9%", "预训练未见该集"),
    ]
    for i, (a, b, c) in enumerate(rows):
        r, col = divmod(i, 3)
        left = Inches(0.4 + col * 4.25)
        top = Inches(1.32 + r * 2.35)
        box(s, left, top, Inches(4.05), Inches(2.15), CARD, STROKE)
        put(s, left + Inches(0.18), top + Inches(0.15), Inches(3.7), Inches(0.45), a, 13, MUTED)
        put(s, left + Inches(0.18), top + Inches(0.65), Inches(3.7), Inches(0.6), b, 20, CYAN, True)
        put(s, left + Inches(0.18), top + Inches(1.35), Inches(3.7), Inches(0.55), c, 13, SOFT)
    footer(s, 13, "结果")

    # 14 OOD + fusion
    s = prs.slides.add_slide(blank)
    bg(s)
    header(s, "4  结果  ·  域外与融合", "没见过的数据集上仍然涨；任意组合都能融，不必出现在预训练对里",
           "域外相对随机猜测：IMU 4.5×、Wi‑Fi 7.6×、mmWave 5×、LiDAR 10.4×。")
    box(s, Inches(0.4), Inches(1.32), Inches(6.2), Inches(2.35), CARD, STROKE)
    put(s, Inches(0.65), Inches(1.45), Inches(5.7), Inches(0.35), "域外单模态相对 SOTA", 14, CYAN, True)
    put(s, Inches(0.65), Inches(1.95), Inches(5.7), Inches(1.45),
        "mmWave  +14.6%\nLiDAR   +13.1%\nWi-Fi    +5.3%",
        16, WHITE)
    box(s, Inches(6.8), Inches(1.32), Inches(6.1), Inches(2.35), CARD, STROKE)
    put(s, Inches(7.05), Inches(1.45), Inches(5.65), Inches(0.35), "域内融合例子", 14, CYAN, True)
    put(s, Inches(7.05), Inches(1.95), Inches(5.65), Inches(1.45),
        "UTD  IMU+Video  33.17%  （高于任一路）\nXRF55  Wi-Fi+mmWave  58.97%\n预训练没有 IMU–Video 对，照样能融。",
        14, SOFT)
    put(s, Inches(0.45), Inches(3.85), Inches(12.4), Inches(0.35), "域外 mRI 融合 vs 多模态基线（表 2）", 13, AMBER, True)
    frows = [
        ("Vision+IMU", "89.6%", "91.7%", "+2.3"),
        ("Vision+mmWave", "58.4%", "64.6%", "+10.7"),
        ("IMU+mmWave", "75.0%", "86.5%", "+13.5"),
        ("三路全部", "85.4%", "92.8%", "+8.6"),
    ]
    for i, (n, b, bb, d) in enumerate(frows):
        left = Inches(0.4 + i * 3.2)
        box(s, left, Inches(4.25), Inches(3.05), Inches(2.05), CARD2, STROKE)
        put(s, left + Inches(0.12), Inches(4.35), Inches(2.8), Inches(0.35), n, 12, MUTED)
        put(s, left + Inches(0.12), Inches(4.75), Inches(2.8), Inches(0.7), bb, 18, GREEN, True)
        put(s, left + Inches(0.12), Inches(5.5), Inches(2.8), Inches(0.6), "基线 " + b + "  (" + d + ")", 11, SOFT)
    footer(s, 14, "结果")

    # 15 vs Cosmo vs MLLM
    s = prs.slides.add_slide(blank)
    bg(s)
    header(s, "4  结果  ·  对照", "同样数据下也强于 Cosmo；对 MLLM 的差距说明「把 IMU 塞进视觉空间」不够",
           "摘要中的「融合最高 +22%、相对 MLLM +25.2%」来自这一组对照。")
    box(s, Inches(0.4), Inches(1.32), Inches(6.2), Inches(5.0), CARD, STROKE)
    put(s, Inches(0.65), Inches(1.48), Inches(5.7), Inches(0.4), "vs Cosmo  ·  UTD IMU–骨架", 15, CYAN, True)
    put(s, Inches(0.65), Inches(2.05), Inches(5.7), Inches(3.95),
        "公平设置：同一对数据、只对齐 2 路、训练 10 次随机种子。\n\n"
        "Cosmo 平均 56.3%\n双模态 Babel  61.46%\n\n"
        "若下游都改成简单 MLP（Cosmo 不再用它那套迭代融合）：\nCosmo(G) 41%  →  Babel 约再高 20 个点。\n\n"
        "把预训练扩到 6 路后，同一 IMU–骨架任务到 63.02%。\n多对齐的模态会回流，帮助没在该任务里出现的组合。",
        13, SOFT)
    box(s, Inches(6.8), Inches(1.32), Inches(6.1), Inches(5.0), CARD, STROKE)
    put(s, Inches(7.05), Inches(1.48), Inches(5.65), Inches(0.4), "vs MLLM  ·  表 4", 15, AMBER, True)
    put(s, Inches(7.05), Inches(2.05), Inches(5.65), Inches(3.95),
        "OneLLM  IMU  6.5%   Video  6.51%\nM4/ImageBind  IMU  5.77%  Video  7.44%\n其余 Wi‑Fi / mmWave / 骨架 / LiDAR：不支持。\n\n"
        "Babel  UTD IMU  31.77%\n       Video 21.35%  Wi‑Fi 33.89%\n       mmWave 50.30%  骨架 61.06%\n       LiDAR 43.91%\n\n"
        "原因：MLLM 的 IMU 基本只在 Ego4D 上对齐，跨到 HAR 设备就塌。",
        13, SOFT)
    footer(s, 15, "结果")

    # 16 Ablation + order + cost
    s = prs.slides.add_slide(blank)
    bg(s)
    header(s, "4  结果  ·  消融与开销", "三块积木都不是可选项；生长顺序几乎不影响终局",
           "这是「创新点」能成立的定量证据。")
    ab = [
        ("去掉预训练塔", "样本太少，训练不收敛", ROSE),
        ("去掉原型网络", "引入新模态后，旧模态相对掉 ~44.7%", ROSE),
        ("去掉自适应权重", "整体最多 -7.2%", ROSE),
        ("换生长顺序", "IMU 波动 <3%，Wi‑Fi <2%", GREEN),
    ]
    for i, (a, b, c) in enumerate(ab):
        left = Inches(0.4 + i * 3.2)
        box(s, left, Inches(1.32), Inches(3.05), Inches(2.45), CARD, c)
        put(s, left + Inches(0.15), Inches(1.48), Inches(2.75), Inches(0.7), a, 14, c, True)
        put(s, left + Inches(0.15), Inches(2.25), Inches(2.75), Inches(1.25), b, 13, SOFT)
    box(s, Inches(0.4), Inches(4.0), Inches(12.5), Inches(2.3), CARD2, STROKE)
    put(s, Inches(0.65), Inches(4.15), Inches(12), Inches(0.35), "系统开销（A100，单样本）", 14, CYAN, True)
    put(s, Inches(0.65), Inches(4.6), Inches(12), Inches(1.45),
        "对齐 MLP 只增加约 8% 推理延迟（IMU：编码器 182.4 ms，对齐头 10.1 ms）。原型网 <1 ms。多模态塔可并行。\n"
        "权重磁盘约 1.1 GB；按选用模态，FP32 显存 1.4–9.92 GB。",
        14, SOFT)
    footer(s, 16, "结果")

    # 17 Case studies
    s = prs.slides.add_slide(blank)
    bg(s)
    header(s, "4  结果  ·  两个案例", "统一空间一旦成立，就能接到扩散模型和 LLM，而不用为传感重训它们",
           "这是「基础模型」叙事的完成件，但论文自己也说仍是初步验证。")
    box(s, Inches(0.4), Inches(1.32), Inches(6.2), Inches(5.0), CARD, STROKE)
    put(s, Inches(0.65), Inches(1.48), Inches(5.7), Inches(0.45), "传感成像  ·  Babel ⊕ unCLIP", 15, CYAN, True)
    put(s, Inches(0.65), Inches(2.1), Inches(5.7), Inches(3.9),
        "把 unCLIP 的图像编码器挂进 Babel，用 L1 把传感 embedding 对齐到图像空间。\n\n"
        "输入：挥手动作的 IMU。\n"
        "环境与画风：文本 prompt。\n"
        "输出：扩散模型生成「看起来在挥手」的图。\n\n"
        "意义：非视觉传感器可以检索视觉表征——无线/惯性「成像」的一条路径。",
        14, SOFT)
    box(s, Inches(6.8), Inches(1.32), Inches(6.1), Inches(5.0), CARD, STROKE)
    put(s, Inches(7.05), Inches(1.48), Inches(5.65), Inches(0.45), "桥接 LLM  ·  Babel ⊕ Video-LLaMA", 15, AMBER, True)
    put(s, Inches(7.05), Inches(2.1), Inches(5.65), Inches(3.9),
        "冻结 Video-LLaMA 视频编码器，让 Babel 所有模态向它对齐（L1）。\n\n"
        "输入挥手 IMU → Babel → Video-LLaMA。\n"
        "LLM 未在 IMU 上特训，仍能区分挥手 vs 下蹲。\n\n"
        "局限（§9）：现在是绕视频空间间接对齐，未来应直接映射到 LLM 原生空间。",
        14, SOFT)
    footer(s, 17, "案例")

    # 18 Takeaway
    s = prs.slides.add_slide(blank)
    bg(s)
    header(s, "5  带走", "Babel 的可迁移结论：物理传感器对齐，先找桥，再冻塔，再管遗忘",
           "代码 aka.ms/project-babel。局限：主任务仍是 HAR；定位/导航等未验证。")
    takes = [
        ("想法", "成对数据与 N 元组是传感多模态的真约束。共享模态当桥，N 路可以拆成二元链。"),
        ("架构", "冻编码器 + 薄对齐头；Trunk/Branch/Junction；共享原型网；梯度加权 + EMA。"),
        ("创新", "可扩展，而不是又一个对比学习。消融证明塔、原型、自适应三者缺一不可。"),
        ("数字", "单模态平均 +12%；融合最高 +22%；IMU 31.77% vs MLLM ~6%；旧模态抗遗忘靠原型网。"),
    ]
    for i, (k, v) in enumerate(takes):
        top = Inches(1.32 + i * 1.18)
        box(s, Inches(0.4), top, Inches(12.5), Inches(1.05), CARD, STROKE)
        put(s, Inches(0.65), top, Inches(1.8), Inches(1.05), k, 16, CYAN, True, valign=MSO_ANCHOR.MIDDLE)
        put(s, Inches(2.6), top, Inches(10.0), Inches(1.05), v, 14, SOFT, valign=MSO_ANCHOR.MIDDLE)
    footer(s, 18, "结论")

    out = r"g:\My Drive\Documents\AI物理感知\Babel论文解读.pptx"
    try:
        prs.save(out)
    except PermissionError:
        out = r"g:\My Drive\Documents\AI物理感知\Babel-SenSys2025-解读.pptx"
        prs.save(out)
    print("SAVED", out)
    print("SLIDES", len(prs.slides))


if __name__ == "__main__":
    build()
