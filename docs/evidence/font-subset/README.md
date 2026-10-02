# Issue 05 字体子集证据

源码基线：`df71c7f20e0c45e9cba9cb78f2d2a91d61413ec9`（PR #12）；05 分支 `codex/font-subset-reuse`，stacked base `codex/document-tools`。测试源码与字体的 SHA-256 在包内 manifest/source-freeze 中记录；字体来源均为固定官方 commit，保留 OFL 原许可。不包含专有系统字体。

| 本次结果 | 范围 |
| --- | --- |
| 功能回归 | 当前全套 464/464；独立 R6 定向 72/72，R7 实际控制字符渲染探针通过 |
| 包消费 | 11 个 `0.1.0-issue05.20261003.10` 本地包，干净目录/缓存验证；包含本票真实字体样例 |
| 自动渲染 | full/subset、native/default 各两页 PDF；全部 Poppler exit 0 且无 stderr；full/subset 与 native/default 逐像素相等 |
| SVG | 8 个页面由真实 Chromium 渲染，FontFaceSet loaded 且 CSS 平台字体 `isCustomFont=true`；保留页面截图和字体记录 |
| PNG 目视 | 检查 full/subset PDF 1–2 页、native/default DOCX→本次 OFD→PDF 1–2 页；subset SVG Chromium 1–2 页，无缺字、裁切、重叠或样式扩散；其他 SVG 截图供复查 |
| macOS Preview | **通过**：2026-10-03 独占时段逐页实际查看 full/subset、native/default 四份 8 页；关闭全部 05 窗口后释放 UI；检查范围及每份 SHA 见 `preview-acceptance.json` |

统一样例（相同 source fonts、两个页面、样式、Unicode 文本和矢量矩形）：

| 载荷 | full bytes | subset bytes | 倍率 / 字形 |
| --- | ---: | ---: | --- |
| LXGW WenKai | 25,575,676 | 963,136 | 26.6x；46867 GID 中保留 140 个轮廓闭包 |
| Noto Sans | 569,208 | 171,292 | 3.32x；3748 GID 中保留 57 个闭包 |
| OFD | 12,998,912 | 397,966 | 32.66x，下降约 96.94% |

四个样式资源绑定两份唯一字体载荷，减少两份重复内容。全量对照也按内容复用，所以体积差主要证明真实子集收益。hmtx 不变；独立直接解析记录 glyf：LXGW `24,450,815 → 31,176`，Noto `401,868 → 6,624`。HarfBuzz 抽样 GID、cluster、advance、offset 与原字体相同。Unicode 文本仍在 OFD/SVG，不是栅格或路径化文本。

产物可从 `evidence.tar.zst` 解开：

```sh
zstd -d evidence.tar.zst -c | tar -xf -
```

包含本次 full/subset OFD 与 PDF、native/default DOCX 衍生 OFD/PDF、subset/native/default SVG、源 DOCX、页面 PNG、环境/哈希/检查记录和测试日志。较大的 full SVG 可由 full OFD 使用公开 SDK 重现，不重复打包。`manifest.json` 记录公开证据内每个文件的 SHA-256 与体积。

DocumentTools 首轮 CI 失败与两轮定位日志保留：空白包会生成没有文字 repertoire 的子集，随后新增水印/注释必须重新绑定原 face，或对明确待注入未知绘制的编辑样例选择 Full。本次修复按此契约执行，不放宽 coverage。再跑 36 样例 / 55 PDF 页通过。

已保留最初样例标题过长及修正后的记录；当前两页页标已拆行，重生成后再次检查。非 BMP P2 的失败轮次和独立 R2 复验亦保留。范围限制详见[字体子集契约](../../font-subset-and-reuse.md)：CFF/变量/未知表/RTL 全量保护；非 BMP PDF 与未建模 CGTransform PDF/SVG 明确失败；已有未知 OFD 资源不裁剪。Preview 仅覆盖列出的 8 页；发布/merge 未授权，不关闭 issue #3。当前 head 的 CI/review 闭环单独记录，不以这些样例推断任意复杂文档保真。

第二轮修复：default-ignorable 控制不要求 cmap 轮廓，PDF 绘制跳过可省略控制符而不改变 OFD 原文或显式 Delta 槽位；方向控制/UVS PDF 明确拒绝。保留 R6 的可见 tofu 失败和 R7 的修复页面。新增 ZWJ/ZWNJ 两页 Preview 待下个独占时段，原 8 页实际验收不受这些无控制符的 guard/绘制分支影响。TTC 容器/face 双预算、OS/2-first 统一样式已有回归；非法 SourceXml 在两模式均拒绝（a6ae520 Full 真实探针为 XmlException/0 字节，并未成功写出）。

R8：绘制时保留已映射 Hangul filler 的真实字宽，普通文本与每 gap Delta 的实测像素/位置一致；全部方向格式控制、蒙古 FVS 语义 PDF 明确拒绝。新增 filler 的无/有 Delta 页面 PNG 复查通过，实际 Preview 与 ZWJ/ZWNJ 一起排解锁后的短时独占；整体视觉门仍保持未完成。

R9：LRM/弃用零宽控制的普通与定位 PDF 成功且无 tofu，九方向格式控制维持明确拒绝。name-only local/host style probe 的物理字节同时注册与缓存 cmap；缺失 filler 省略，映射 filler 保留宽度。独立实际 name-only 24 项定向验证与像素对比通过；探针使用公开 `PdfFontRegistry.CreateResolver(host)` 组合约定。记录区分不符合约定的旧探针，不将它冒充成功证据。
