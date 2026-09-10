# TimesFM-3: A zero-shot foundation model for multivariate forecasting

> 原文：Google Research Blog，2026-08-31
> 作者：Ayush Jain、Rajat Sen（Google Research）
> 链接：https://research.google/blog/timesfm-3-a-zero-shot-foundation-model-for-multivariate-forecasting/
> 致谢：Yichen Zhou、Petros Mol、Abhimanyu Das、Samet Oymak

TimesFM-3 是 Google Research 的时序基础模型（TSFM）第三代，可在**单次前向传播**内完成高精度**多变量**时间序列预测，在主要基准上显著超过其他预测模型。

## 背景

自 2024 年 TimesFM 发布以来，时序基础模型已在零售、金融、可观测性、制造、医疗、自然科学等领域落地。但直到 TimesFM-2.5（2025-09），模型只支持**单变量**预测：只能用单条序列的历史。真实问题本质上是多变量的——例如预测冰淇淋销量，不能只看历史销量，还要看相关产品（蛋筒、糖浆）的销量、历史客流、以及已知未来事件（天气预报、促销、节假日）。

## 模型规模与能力

- **330M 参数**，预训练语料超过 **1 万亿个时间点**（真实 + 合成）。
- 原生多变量、零样本（zero-shot）：联合预测多条共同演化的序列，无需任务特定微调。
- 原生支持：
  - **多目标（Multiple targets）**：同时预测多条相关序列，支持点预测与分位数预测。
  - **历史协变量（Past covariates）**：只有历史已知的特征（如历史客流）。
  - **历史-未来协变量（Past-future / dynamic covariates）**：未来已知的特征（如促销计划、天气预报）。

## 架构与推理

基于前几代的 **decoder-only transformer**：

- **Patch**：连续 32 个时间点为一组（patch=32）；每条序列单独做归一化（同 TimesFM-2.5），以兼容量纲差异巨大的序列。
- **多变量 token 构造**：目标序列与历史协变量的 token 直接由单个 patch 构成；历史-未来协变量采用 **lookahead** 策略——每个 token 拼接当前 patch 与未来 patch，让模型看到已知的未来信号。
- **交替注意力（Alternating attention）**：token 经过输入残差块后进入主干，主干是一个 2D 网格：
  1. **因果时间注意力（Causal temporal attention）**：沿时间轴横向注意力，严格因果，防止数据泄漏——token 只能看自己序列的过去。
  2. **全变量注意力（Full variate attention）**：沿变量轴纵向注意力——任一时刻，token 可以看到数据集中所有其他序列，学习跨序列相关（如促销如何影响另一条的销量）。
  两种注意力交替堆叠多层，把时间模式与跨序列关系融合。

## 非自回归解码：单次前向出整段预测

此前版本逐 patch 自回归生成，带来延迟、误差累积与算力开销。TimesFM-3 用 **Contiguous Patch Masking（连续 patch 掩码）**：在观测上下文后追加未来 horizon 的掩码占位 token——目标与历史协变量在 horizon 段被掩码（未来未知），历史-未来协变量保持可见（提供节假日、计划事件等已知未来信号）。通过交替注意力层，模型**同时填出所有被掩码的 horizon patch，无需迭代循环**。每个目标序列在每个 horizon 步输出 **9 个分位数（10%–90%）**，给出完整的概率预测视图。

## 示例：促销计划

预测下月冰淇淋销量。单变量模型只能把历史周模式外推，不知道未来哪天有促销；TimesFM-3 多变量模式把「促销计划」作为历史-未来协变量输入，从历史上下文中学到促销与销量提升的关系，并应用到未来促销日——预测曲线在每个促销日抬高约 20%。

## 评测

在三个公开基准 **GIFT-Eval、FEV-Bench、TIME** 上，TimesFM-3 在点预测与概率预测两类指标上均为**预训练基础模型第一名**（对比 Chronos-2、Toto 2.0 系列、TimesFM-2.5 等）：

- **单变量模式**（不用任何协变量与跨序列信息）已匹配或超过其他可复现模型。
- **多变量模式**进一步提升，取得两项指标的最佳平均排名。

具体：fev-bench 100 个真实任务总排名第一；TIME 基准 50 个领域数据集 / 98 个评测任务总排名第一；GIFT-Eval 基础模型中排名第一。

## 可用性与许可

- 代码：GitHub `google-research/timesfm`（Apache-2.0）。
- 权重：Hugging Face `google/timesfm-3.0-pytorch`，**TimesFM Non-Commercial License v1.0，仅限非商用、非生产用途**（2.5 及之前版本权重为 Apache-2.0）。
- BigQuery `AI.FORECAST` 集成将陆续上线。
