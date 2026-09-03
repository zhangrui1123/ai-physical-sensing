# -*- coding: utf-8 -*-
"""AI 物理感知平台 PPT：一页一事，标题即结论。"""
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
WHITE = RGBColor(0xF8, 0xFA, 0xFC)
MUTED = RGBColor(0x94, 0xA3, 0xB8)
SOFT = RGBColor(0xCB, 0xD5, 0xE1)
FONT = "Microsoft YaHei"
TOTAL = 12


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
    lines = str(text).split("\n")
    for i, line in enumerate(lines):
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
        slide, Inches(0.4), Inches(7.28), Inches(9.2), Inches(0.22),
        ("AI 物理感知  ·  " + section) if section else "AI 物理感知平台",
        9, MUTED, valign=MSO_ANCHOR.MIDDLE,
    )
    put(
        slide, Inches(11.3), Inches(7.28), Inches(1.6), Inches(0.22),
        f"{page}  /  {TOTAL}", 9, MUTED, align=PP_ALIGN.RIGHT, valign=MSO_ANCHOR.MIDDLE,
    )


def header(slide, kicker, title, subtitle=None):
    rect(slide, 0, 0, Inches(0.12), H, CYAN)
    put(slide, Inches(0.45), Inches(0.14), Inches(12.4), Inches(0.26), kicker, 11, CYAN, True)
    put(slide, Inches(0.45), Inches(0.38), Inches(12.5), Inches(0.52), title, 22, WHITE, True)
    if subtitle:
        put(slide, Inches(0.45), Inches(0.88), Inches(12.5), Inches(0.34), subtitle, 13, MUTED)


def so_what(slide, text):
    box(slide, Inches(0.4), Inches(6.48), Inches(12.5), Inches(0.72), CARD2, AMBER)
    rect(slide, Inches(0.4), Inches(6.48), Inches(0.1), Inches(0.72), AMBER)
    put(slide, Inches(0.7), Inches(6.48), Inches(12.0), Inches(0.72),
        "所以  ·  " + text, 13, WHITE, True, valign=MSO_ANCHOR.MIDDLE)


def build():
    prs = Presentation()
    prs.slide_width, prs.slide_height = W, H
    blank = prs.slide_layouts[6]

    # ========== 1 封面 ==========
    s = prs.slides.add_slide(blank)
    bg(s)
    rect(s, 0, 0, Inches(0.16), H, CYAN)
    put(s, Inches(0.7), Inches(1.25), Inches(12), Inches(0.32),
        "立项汇报  ·  物理 = 视频 / 音频 / 文字之外的传感器信号", 14, AMBER, True)
    put(s, Inches(0.7), Inches(1.72), Inches(12), Inches(0.85), "AI 物理感知平台", 40, WHITE, True)
    put(s, Inches(0.7), Inches(2.58), Inches(12), Inches(0.42),
        "做信号操作系统，不做语言大模型，不做世界视频", 18, SOFT)
    box(s, Inches(0.7), Inches(3.2), Inches(11.9), Inches(1.45), CARD, STROKE)
    put(s, Inches(0.95), Inches(3.35), Inches(11.4), Inches(1.15),
        "VLM 已吃掉图声文；雷达 / IMU / EMG / UWB 仍是一路一个小模型。我们补中间层。",
        16, SOFT)
    for i, (k, v) in enumerate([
        ("做什么", "四座塔：雷达 · IMU · EMG · UWB"),
        ("怎么验证", "加进来必须加分"),
        ("先打哪", "无接触生命体征"),
    ]):
        left = Inches(0.7 + i * 4.05)
        box(s, left, Inches(4.9), Inches(3.85), Inches(1.55), CARD2, STROKE)
        put(s, left + Inches(0.2), Inches(5.05), Inches(3.45), Inches(0.32), k, 13, CYAN, True)
        put(s, left + Inches(0.2), Inches(5.42), Inches(3.45), Inches(0.85), v, 15, WHITE)
    put(s, Inches(0.7), Inches(6.7), Inches(8), Inches(0.3), "2026  ·  蓝图", 13, MUTED)
    put(s, Inches(10.3), Inches(6.7), Inches(2.5), Inches(0.3), f"1  /  {TOTAL}", 13, MUTED, align=PP_ALIGN.RIGHT)

    # ========== 2 请拍板 ==========
    s = prs.slides.add_slide(blank)
    bg(s)
    header(s, "0  请拍板", "立项只拍三件事，其余都是证据",
           "讲完这页，后面 10 分钟只为这三句补材料。")
    asks = [
        (CYAN, "1  做什么",
         "信号操作系统。\n雷达 / IMU / EMG / UWB 进同一套时钟、tokenizer、评测。",
         "不做语言大模型、世界视频、芯片、整机、眼镜整机。"),
        (AMBER, "2  怎么验证",
         "加上雷达、IMU、EMG 或 UWB，任务分必须上升。\n视觉解锁率、FID 一律不算。",
         "加进来不加分，禁止称「多传感融合」。"),
        (GREEN, "3  先打哪",
         "无接触生命体征。\n不贴电极，被子里也要有呼吸。",
         "90 天演示「加雷达，被子里才有读数」。钥匙、写字、运动随后建塔。"),
    ]
    for i, (c, t, a, b) in enumerate(asks):
        left = Inches(0.4 + i * 4.3)
        box(s, left, Inches(1.35), Inches(4.15), Inches(4.95), CARD, STROKE)
        rect(s, left, Inches(1.35), Inches(0.12), Inches(4.95), c)
        put(s, left + Inches(0.3), Inches(1.55), Inches(3.65), Inches(0.45), t, 18, c, True)
        put(s, left + Inches(0.3), Inches(2.15), Inches(3.65), Inches(2.2), a, 16, WHITE)
        put(s, left + Inches(0.3), Inches(4.5), Inches(3.65), Inches(1.5), b, 14, SOFT)
    footer(s, 2, "拍板")

    # ========== 3 规格 ==========
    s = prs.slides.add_slide(blank)
    bg(s)
    header(s, "1  规格", "主吃六路；视频 / 音频 / 文本只作对齐教师",
           "禁止把摄像头当训练与评测的主输入。")
    defs = [
        (GREEN, "本平台取", "雷达 · IMU · EMG · 接近光 · TOF · TP", "主模态"),
        (MUTED, "不做 A", "世界模型：Cosmos / Dreamer / JEPA", "不做"),
        (MUTED, "不做 B", "Occupancy：Tesla / World Labs Marble", "不做"),
    ]
    for i, (color, tag, desc, badge) in enumerate(defs):
        left = Inches(0.4 + i * 4.3)
        box(s, left, Inches(1.32), Inches(4.15), Inches(0.95), CARD, color)
        put(s, left + Inches(0.16), Inches(1.38), Inches(2.4), Inches(0.28), tag, 12, color, True)
        put(s, left + Inches(2.4), Inches(1.38), Inches(1.55), Inches(0.28), badge, 11, color, True, PP_ALIGN.RIGHT)
        put(s, left + Inches(0.16), Inches(1.7), Inches(3.8), Inches(0.48), desc, 14, WHITE)
    fams = [
        ("雷达", "距离 / 速度 / 微动", "第一塔：生命体征"),
        ("IMU", "加速度 / 角速度", "第二塔：运动状态"),
        ("EMG", "肌腱表面电位", "第三塔：微手势、写字"),
        ("UWB", "无线电 TOF", "第四塔：数字钥匙"),
        ("接近光 / TOF", "近距有无 / 距离图", "常开、暗光可用"),
        ("TP", "触点 / 掌误触", "接触事件"),
    ]
    for i, (a, b, c) in enumerate(fams):
        r, col = divmod(i, 3)
        left = Inches(0.4 + col * 4.3)
        top = Inches(2.42 + r * 1.95)
        box(s, left, top, Inches(4.15), Inches(1.82), CARD, STROKE)
        put(s, left + Inches(0.18), top + Inches(0.14), Inches(3.8), Inches(0.42), a, 18, CYAN, True)
        put(s, left + Inches(0.18), top + Inches(0.62), Inches(3.8), Inches(0.45), b, 15, WHITE)
        put(s, left + Inches(0.18), top + Inches(1.14), Inches(3.8), Inches(0.5), c, 14, AMBER)
    so_what(s, "四座塔按商用场景立项。义肢、针电极不做。")
    footer(s, 3, "规格")

    # ========== 4 为何现在 ==========
    s = prs.slides.add_slide(blank)
    bg(s)
    header(s, "1  为什么现在", "图声文有基础模型，物理信号还没有操作系统",
           "四条公开结果。汇报时只念标题，细节留问答。")
    facts = [
        ("雷达能进大模型", "HoloLLM：mmWave 注入 VLM，人体感知最高 +30%。"),
        ("EMG 已开卖，仍要试戴", "Meta Neural Band 2025-09 开卖，近 20 万人训练。"),
        ("IMU 必须自有塔", "Babel：ImageBind 对 IMU 几乎无效。"),
        ("EMG / 雷达同样", "LIMU-BERT 已上移动端。"),
    ]
    for i, (t, d) in enumerate(facts):
        r, c = divmod(i, 2)
        left = Inches(0.4 + c * 6.45)
        top = Inches(1.35 + r * 2.15)
        box(s, left, top, Inches(6.2), Inches(2.0), CARD, STROKE)
        put(s, left + Inches(0.22), top + Inches(0.2), Inches(5.75), Inches(0.5), t, 18, CYAN, True)
        put(s, left + Inches(0.22), top + Inches(0.8), Inches(5.75), Inches(1.0), d, 15, SOFT)
    so_what(s, "缺口在接入与表征，不在再训一个视觉或语言大模型。")
    footer(s, 4, "缺口")

    # ========== 5 代际 ==========
    s = prs.slides.add_slide(blank)
    bg(s)
    header(s, "2  代际", "主线不是模型更大，而是信号如何被编码",
           "G2 交付，G3 做底座，主攻 G4。Hub 已出货，缺的是分传感器 token。")
    gens = [
        ("G0", "滤波", "Kalman / 阈值", MUTED),
        ("G1", "单模态深度", "IMU-LSTM / 雷达 CNN", SOFT),
        ("G2", "Hub 规则+小网", "端侧量产主力", CYAN),
        ("G3", "传感基础模型", "LIMU-BERT / Babel", INDIGO),
        ("G4", "分传感器 token", "当前窗口", AMBER),
        ("G5", "传感 OS", "预埋接口", GREEN),
    ]
    for i, (g, name, tech, color) in enumerate(gens):
        left = Inches(0.32 + i * 2.16)
        hot = g in ("G3", "G4")
        box(s, left, Inches(1.32), Inches(2.06), Inches(2.05),
            RGBColor(0x1C, 0x2C, 0x48) if hot else CARD,
            AMBER if g == "G4" else (CYAN if g == "G3" else STROKE))
        put(s, left + Inches(0.08), Inches(1.42), Inches(1.9), Inches(0.38), g, 18,
            AMBER if g == "G4" else (CYAN if g == "G3" else MUTED), True, PP_ALIGN.CENTER)
        put(s, left + Inches(0.08), Inches(1.88), Inches(1.9), Inches(0.6), name, 14, WHITE, True, PP_ALIGN.CENTER)
        put(s, left + Inches(0.08), Inches(2.55), Inches(1.9), Inches(0.65), tech, 12, SOFT, align=PP_ALIGN.CENTER)
    jumps = [
        ("不跳 G0→G1", "换机型，手工阈值全失效。LIMU-BERT 证明 IMU 可离开阈值。", CYAN),
        ("不跳 G1→G2", "入袋 / 抬腕各写 if-else。Hub 已把多路接到低功耗核。", INDIGO),
        ("不跳 G3→G4", "每加一路雷达、IMU、UWB 或 EMG 重写产线。Babel 证明可以只加一座塔。", AMBER),
    ]
    for i, (k, v, c) in enumerate(jumps):
        top = Inches(3.55 + i * 0.9)
        box(s, Inches(0.4), top, Inches(12.5), Inches(0.82), CARD, STROKE)
        rect(s, Inches(0.4), top, Inches(0.1), Inches(0.82), c)
        put(s, Inches(0.7), top, Inches(3.2), Inches(0.82), k, 15, c, True, valign=MSO_ANCHOR.MIDDLE)
        put(s, Inches(4.0), top, Inches(8.6), Inches(0.82), v, 15, SOFT, valign=MSO_ANCHOR.MIDDLE)
    footer(s, 5, "代际")

    # ========== 6 验证 ==========
    s = prs.slides.add_slide(blank)
    bg(s)
    header(s, "3  验证", "加上哪座塔，哪个场景就必须更好",
           "无采样率、单位、时间基的数据包，Ingest 直接拒收。")
    rows = [
        ("雷达", "生命体征", "被子里没有读数", "被子里也有呼吸"),
        ("IMU", "运动状态", "抬腕 / 入袋分不清", "抬腕 / 入袋可识别"),
        ("EMG", "眼镜写字", "微点 / 笔画做不到", "微点 / 笔画可用"),
        ("UWB", "数字钥匙", "BLE 会被中继骗过", "抗中继，走近才开"),
    ]
    box(s, Inches(0.4), Inches(1.32), Inches(12.5), Inches(0.5), RGBColor(0x1A, 0x2A, 0x42))
    for j, h in enumerate(["塔", "场景", "没有", "加上"]):
        put(s, Inches(0.6 + j * 3.1), Inches(1.32), Inches(2.9), Inches(0.5), h, 14, CYAN, True, valign=MSO_ANCHOR.MIDDLE)
    for i, (a, b, c, d) in enumerate(rows):
        top = Inches(1.88 + i * 1.05)
        box(s, Inches(0.4), top, Inches(12.5), Inches(0.95), CARD if i % 2 == 0 else CARD2, STROKE)
        put(s, Inches(0.6), top, Inches(2.9), Inches(0.95), a, 16, WHITE, True, valign=MSO_ANCHOR.MIDDLE)
        put(s, Inches(3.7), top, Inches(2.9), Inches(0.95), b, 15, AMBER, valign=MSO_ANCHOR.MIDDLE)
        put(s, Inches(6.8), top, Inches(2.9), Inches(0.95), c, 15, MUTED, valign=MSO_ANCHOR.MIDDLE)
        put(s, Inches(9.9), top, Inches(2.9), Inches(0.95), d, 15, WHITE, valign=MSO_ANCHOR.MIDDLE)
    footer(s, 6, "验证")

    # ========== 7 对标 ==========
    s = prs.slides.add_slide(blank)
    bg(s)
    header(s, "4  对标", "缺的不是器件，是中间那层操作系统",
           "抄时钟与 schema，抄分传感器塔。不抄封闭小模型，不卖整机，不造芯片。")
    layers3 = [
        ("上  整机", "不卖整机",
         "Meta Neural Band、BMW 数字钥匙、睡眠雷达卖体验。责任绑在整机，原始波形不对外开放。", CYAN, False),
        ("中  信号 OS  ← 我们", "主航道",
         "Hub / Core Motion 停在 G2。我们补四座 tokenizer，以及「加雷达 / 加 IMU / 加 EMG / 加 UWB 必须更好」的表。", AMBER, True),
        ("下  器件", "不造芯片",
         "Bosch IMU、Infineon 雷达、ST UWB、干电极 EMG。集成探头，不自研硅。", MUTED, False),
    ]
    for i, (t, badge, a, color, hot) in enumerate(layers3):
        top = Inches(1.35 + i * 1.7)
        box(s, Inches(0.4), top, Inches(12.5), Inches(1.55),
            RGBColor(0x1C, 0x2C, 0x48) if hot else CARD, color if hot else STROKE)
        put(s, Inches(0.7), top + Inches(0.18), Inches(8.5), Inches(0.4), t, 20, color, True)
        put(s, Inches(9.3), top + Inches(0.18), Inches(3.3), Inches(0.4), badge, 14,
            GREEN if hot else MUTED, True, PP_ALIGN.RIGHT)
        put(s, Inches(0.7), top + Inches(0.7), Inches(12.0), Inches(0.65), a, 16, SOFT)
    footer(s, 7, "对标")

    # ========== 8 四场景对照 ==========
    s = prs.slides.add_slide(blank)
    bg(s)
    header(s, "5  商用场景", "四件事倒逼四座塔：雷达 · IMU · EMG · UWB",
           "汇报只讲这一页。先打中间这一列。细证据在附录，被问再翻。")
    scenes = [
        ("眼镜写字 / 微手势",
         ["手在桌下也能写、口袋里也能点", "眼镜 OEM / XR", "手在桌下、口袋里，会议室不能出声", "加 EMG：微点 / 写字可用", "EMG 塔 + 对比表"]),
        ("无接触生命体征  ← 先打",
         ["不贴电极，被子里也要有呼吸心跳", "养老 / 睡眠硬件", "卧室卫生间不能常开录像", "加雷达：被子里也有读数", "雷达微动塔 + 对比表"]),
        ("运动状态识别",
         ["抬腕亮屏、入袋静音、跌倒报警", "手机 / 手表 / 耳机", "摄像头常开功耗高、隐私差", "加 IMU：抬腕 / 入袋可识别", "IMU 塔 + 对比表"]),
        ("UWB 数字钥匙",
         ["走近要开，假的「很近」要识破", "车身电子 + 手机 OEM", "中继是无线电欺骗，不是图像", "加 UWB：抗中继，走近才开", "UWB 塔 + 抗中继评测"]),
    ]
    box(s, Inches(0.35), Inches(1.3), Inches(12.6), Inches(0.48), RGBColor(0x1A, 0x2A, 0x42))
    put(s, Inches(0.45), Inches(1.3), Inches(2.3), Inches(0.48), "", 12, CYAN, True, valign=MSO_ANCHOR.MIDDLE)
    for j, (name, _) in enumerate(scenes):
        put(s, Inches(2.8 + j * 2.55), Inches(1.3), Inches(2.4), Inches(0.48), name, 13, CYAN, True,
            valign=MSO_ANCHOR.MIDDLE)
    labels = ["一句话", "谁付钱", "为何不用摄像头", "加哪路更好", "我们交什么"]
    for i, lab in enumerate(labels):
        top = Inches(1.84 + i * 0.88)
        box(s, Inches(0.35), top, Inches(12.6), Inches(0.82), CARD if i % 2 == 0 else CARD2, STROKE)
        put(s, Inches(0.45), top, Inches(2.25), Inches(0.82), lab, 13, AMBER, True, valign=MSO_ANCHOR.MIDDLE)
        for j, (_, cells) in enumerate(scenes):
            put(s, Inches(2.8 + j * 2.55), top, Inches(2.4), Inches(0.82), cells[i], 13, WHITE,
                valign=MSO_ANCHOR.MIDDLE)
    footer(s, 8, "场景")

    # ========== 9 先打生命体征 ==========
    s = prs.slides.add_slide(blank)
    bg(s)
    header(s, "5  先打哪", "无接触生命体征：90 天证明「加雷达，被子里才有读数」",
           "不先做眼镜整机，也不先做开锁产品。先立能当场验证的雷达微动塔。")
    box(s, Inches(0.4), Inches(1.32), Inches(8.0), Inches(5.0), CARD, STROKE)
    put(s, Inches(0.65), Inches(1.48), Inches(7.5), Inches(0.4), "为什么先打这一件，不先打另外三件", 16, CYAN, True)
    triples = [
        ("体征", "卧室不能装摄像头；静息就能录；加雷达当场能看。"),
        ("运动", "后做。抬腕 / 入袋已有 Hub，差异化小。"),
        ("眼镜", "后做。干电极要拟合，Meta 自己还强制店内试戴。"),
        ("钥匙", "后做。CCC / 主机厂闭环，外面难录真实中继包。"),
    ]
    for i, (a, b) in enumerate(triples):
        top = Inches(2.05 + i * 0.75)
        put(s, Inches(0.75), top, Inches(1.6), Inches(0.7), a, 16, AMBER, True, valign=MSO_ANCHOR.MIDDLE)
        put(s, Inches(2.4), top, Inches(5.7), Inches(0.7), b, 16, WHITE, valign=MSO_ANCHOR.MIDDLE)
    put(s, Inches(0.65), Inches(4.7), Inches(7.5), Inches(1.35),
        "90 天成功：盖毯静息呼吸可出数；加雷达后读数稳定。\n"
        "失败：加雷达前后分数几乎不变——说明雷达没起作用。",
        15, SOFT)
    box(s, Inches(8.6), Inches(1.32), Inches(4.3), Inches(5.0), CARD2, STROKE)
    put(s, Inches(8.8), Inches(1.48), Inches(3.95), Inches(0.4), "90 天交什么", 16, AMBER, True)
    put(s, Inches(8.8), Inches(2.05), Inches(3.95), Inches(4.0),
        "冻结：雷达微动 schema\n"
        "交付：录包 → 同步 → Eval\n"
        "演示：加雷达，被子里才有读数\n\n"
        "随后三座塔\n"
        "IMU 运动状态\n"
        "UWB 数字钥匙\n"
        "EMG 微手势 / 桌面写字\n\n"
        "不交\n"
        "医疗级心率、眼镜整机、开锁产品、语言头",
        15, SOFT)
    footer(s, 9, "第一包")

    # ========== 10 路线 ==========
    s = prs.slides.add_slide(blank)
    bg(s)
    header(s, "6  怎么做", "每个阶段必须有可证伪的成功标准",
           "不做：语言大模型、世界视频、整机 UX、自研芯片、医疗注册。")
    phases = [
        ("P0  0–3 月", "雷达体征塔",
         "冻结雷达微动 schema。\n成功：加雷达，被子里有读数。\n失败：仍用 CSV 对时钟。", GREEN),
        ("P1  3–6 月", "IMU 运动塔",
         "抬腕 / 入袋 / 跌倒。\n成功：加 IMU 必须更好。\n失败：复用 Hub 规则。", CYAN),
        ("P2  6–12 月", "UWB 钥匙塔",
         "数字钥匙 Pack。\n成功：只加一座塔，不改主干。\n失败：每加一路分叉仓库。", INDIGO),
        ("P3  12–18 月", "EMG 写字塔",
         "微手势 + 桌面写字。\n成功：加 EMG 必须更好。\n出汗 / 位移进评测。", AMBER),
    ]
    for i, (a, b, c, color) in enumerate(phases):
        left = Inches(0.35 + i * 3.24)
        box(s, left, Inches(1.35), Inches(3.1), Inches(5.7), CARD, STROKE)
        rect(s, left, Inches(1.35), Inches(3.1), Inches(0.1), color)
        put(s, left + Inches(0.18), Inches(1.6), Inches(2.75), Inches(0.45), a, 16, color, True)
        put(s, left + Inches(0.18), Inches(2.15), Inches(2.75), Inches(0.45), b, 18, WHITE, True)
        put(s, left + Inches(0.18), Inches(2.8), Inches(2.75), Inches(3.8), c, 15, SOFT)
    footer(s, 10, "路线")

    # ========== 11 决策 ==========
    s = prs.slides.add_slide(blank)
    bg(s)
    header(s, "决策", "请拍板的就是封面那三句",
           "第一包：无接触生命体征。第二包：IMU 运动状态。第三包：UWB 数字钥匙。第四包：EMG 写字。")
    lines = [
        ("规格", "雷达 / IMU / EMG / 接近光 / TOF / TP。UWB 按无线电 TOF 接入。"),
        ("不做", "语言大模型、世界视频、芯片、整机、眼镜整机、医疗级心率注册。"),
        ("验证", "加雷达 / 加 IMU / 加 EMG / 加 UWB；不加分即失败。"),
        ("先打", "无接触生命体征：90 天交出「加雷达，被子里才有读数」。"),
        ("场景", "眼镜写字 / 微手势 · 无接触生命体征 · 运动状态 · UWB 数字钥匙。"),
        ("产品", "Ingest + 四座 Tokenizer + Eval。"),
    ]
    for i, (k, v) in enumerate(lines):
        top = Inches(1.32 + i * 0.85)
        box(s, Inches(0.4), top, Inches(12.5), Inches(0.76), CARD, STROKE)
        put(s, Inches(0.65), top, Inches(1.8), Inches(0.76), k, 18, CYAN, True, valign=MSO_ANCHOR.MIDDLE)
        put(s, Inches(2.6), top, Inches(10.0), Inches(0.76), v, 16, WHITE, valign=MSO_ANCHOR.MIDDLE)
    footer(s, 11, "决策")

    # ========== 12 来源 ==========
    s = prs.slides.add_slide(blank)
    bg(s)
    header(s, "附录", "被问到数字时翻这一页，汇报正文不念",
           "厂商口径单独标明。阈值用自有 Pack 标定后写入 Eval。")
    srcs = [
        "Meta Neural Band（2025-09）：与 Ray-Ban Display 捆绑 799 美元；近 20 万受试者；强制店内试戴；手写消息未上线",
        "Mudra Link 约 249 美元；Apple 表带电极专利 + EMBridge（NeurIPS 2025 workshop），手表未出货",
        "AKM AK5816AIM（2026-07 量产）：无摄像头跌倒 + 呼吸；模块口径写静息 / 睡眠，走动伪迹未收口",
        "BMW Digital Key Plus / 奔驰：UWB 约 4 m 定左右侧防中继；CCC Digital Key 已上车",
        "ST ST64UWB（2026，厂商口径）：2025 年 UWB 器件 5.27 亿颗",
        "Babel SenSys 2025：ImageBind 对 IMU 几乎无效；LIMU-BERT 已上移动端；HoloLLM mmWave +30%",
        "观察不进本版 TOP4：座舱 CPD、房间人感、肌电义肢（器械渠道）",
    ]
    for i, t in enumerate(srcs):
        put(s, Inches(0.55), Inches(1.38 + i * 0.7), Inches(12.2), Inches(0.65), "·  " + t, 14, SOFT)
    footer(s, 12, "来源")

    out = r"g:\My Drive\Documents\AI物理感知\AI物理感知平台-传感器口径-汇报.pptx"
    try:
        prs.save(out)
    except PermissionError:
        out = r"g:\My Drive\Documents\AI物理感知\AI物理感知平台-传感器口径-汇报-v2.pptx"
        prs.save(out)
    print("SAVED", out)
    print("SLIDES", len(prs.slides))


if __name__ == "__main__":
    build()
