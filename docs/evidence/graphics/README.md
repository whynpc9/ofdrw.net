# 04 原生绘图验收证据

R5 运行源码冻结于 `dc5efcfc506b3eb91d2a3b80fa54fdeed8186231`，基线为 PR #12 的 `df71c7f20e0c45e9cba9cb78f2d2a91d61413ec9`。[PR #13](https://github.com/whynpc9/ofdrw.net/pull/13) 的 stacked base 为 `codex/document-tools`。本次候选 `5dd3dd7` 对应的六组 PDF 共 **10 页已于 2026-10-03 在 macOS Preview 实际逐页验收通过**，六个任务 PDF 窗口已核对路径并关闭，GUI 已释放。源码、样例、字体和页面字节没有修改；最终文档 head 外部复核仍需读回。

状态与逐文件结果见 [acceptance.json](acceptance.json)，文件尺寸/SHA-256/环境/模式见 [manifest.json](manifest.json)。功能、自动渲染、PNG 与 Preview 分别记录，旧 R1 验收没有用于替代本轮。

| 检查层 | R5 本次范围 | 结果 |
| --- | --- | --- |
| 功能回归 | 全套 429 项；独立新增 29 项原生往返、矩阵/状态/clip、字体身份、数字/强调、预算/取消与栅格回归 | 通过 |
| API 文档 | Graphics 新公开 API CS1591 | 0 新警告 |
| 本地包消费 | 11 个 `0.1.0-graphics.20261003.r5final` 包实际重建、安装和转换，直接从包调用公开 Graphics | 通过；未发布 |
| 自动渲染 | 图形/往返/Native/default 各两页，name-only/分数 clip 各一页；PDF 10 页、SVG 10 页 | 通过；含尖角及分数边框实际像素探针 |
| PNG 视觉 | 20 页 PDF/SVG PNG；本轮重新生成并核对与 R3/R4 页图哈希 | 检查范围内通过 |
| macOS Preview | 本次六组 OFD→PDF 共 10 页，准确路径与 PDF 哈希核对，单页显示逐页检查 | 通过；六个 PDF 窗口关闭，GUI 已释放 |
| 最新 head 外部审核 | 候选 `5dd3dd7` 六检查全绿；Preview P1 已完成实际要求，最终记录 commit 待双 bot/CI | 待最终文档 head 读回 |

## 实际 Preview 范围

- `graphics.pdf` 1–2：中英票面、比例字距、局部粗体/斜体/颜色、基线、旋转框字、页脚；非等比线宽、曲线、miter 尖角、EvenOdd 镂空/NonZero 实心、冻结裁剪及恢复控制对象。
- `graphics-roundtrip.pdf` 1–2：以上原生对象往返后完整保留，无重复文本或视觉补绘重影。
- `graphics-name-only.pdf` 1：正体控制、原生 faux 与用户 CTM 的单次强调、固定基线和裁剪。OFD 为 name-only 声明，导出 PDF 明确注册相同 OFL 字体，不借用专有系统字库。
- `graphics-fractional-clip.pdf` 1：100000 mm 逻辑框在 0.9996 变换及负平移下，末端蓝色右边框在 130 mm 辅助线左侧保持可见；页外左部裁切是样例明确预期。
- `baseline-native.pdf` 1–2、`baseline-default.pdf` 1–2：本次许可字体变体 `generated-layout.docx`→显式 Native/default OFD→PDF；原文、局部样式、表格填充/边框、对齐和确定的两页分页正常。没有用直接 DOCX→PDF 代替。

本轮未见中文缺字/乱码、英文异常字距、样式扩散、重叠/重影、非预期裁切或额外空白页。只对列出的实际样例和页面作结论，不推断任意复杂 Word、通用 shaping 或其他任务产物均保真。全部打开路径位于本 worktree 的 `artifacts/graphics/review5-final/`；每份 PDF 的 SHA-256 与清单一致。

## 实现及证据边界

公开样例只引用 Layout（传递 Core/Packaging），两页包内 18 个 PathObject、30 个 TextObject、1 份字体、0 个 ImageObject。name-only 一页为 7 个 TextObject/4 个 PathObject、不含 FontFile 的声明。04 没有增加 flow/table/Canvas/Skia 适配或第二套字体服务；05 负责字体子集、复用与缺字策略，04 单票不宣称 issue #3 全部完成。

数值使用有界 BigInteger 精确奇异判定和统一保真普通十进制 CTM；保留 Text SourceXml 所有权及 Box/Size/LineWidth 原三位策略。现有 Core 强调服务共用 faux 组合，绘制前原子预检；advances/path/clip XML 在剩余预算内逐段格式化，线宽边界使用实际写出值。参数范围、极端数值限制与资源所有权见 [设计契约](../../graphics-design-contract.md)。

所有公开字体载荷仅为固定 installer 的 Noto Sans CJK SC Regular，并附 OFL；SVG 使用隔离许可字体目录。完整字体 OFD 约 11.6 MB，name-only 约 2 KB 并依赖阅读器环境；未把字体子集计入 04 功能。

`evidence.tar.zst` 保存本次 OFD/PDF/SVG、20 页 PNG、Unicode 原文、许可 DOCX、字体/OFL、完整日志和清单。全部 **63 个载荷文件已默认 zstd 解码器全新提取并校验尺寸/SHA-256**。归档采用 128 MiB 窗口；可执行 `zstd -d -c docs/evidence/graphics/evidence.tar.zst | tar -xf - -C artifacts/graphics-review` 后按 manifest 校验。历史完整归档与失败探针保留在 Git 提交及当前日志中；测试总数为 R1 410、R2 417、R3 424、R4 425、R5 429。

复现执行 `scripts/run-graphics-e2e.sh artifacts/graphics/review5-final`，字体由固定 installer 安装到 `artifacts/graphics-fonts`。Poppler PDF 110 DPI，librsvg SVG 白背景 720px，严格采用 AGENTS 的 CLI_HOME/skip/telemetry/单节点 flags。重新生成后须再次实际 Preview，不因 PNG 相同省略此门。
