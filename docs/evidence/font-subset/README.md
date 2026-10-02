# Issue 05 字体子集证据

源码基线：`df71c7f20e0c45e9cba9cb78f2d2a91d61413ec9`（PR #12）；05 分支 `codex/font-subset-reuse`，stacked base `codex/document-tools`。测试源码与字体的 SHA-256 在包内 manifest/source-freeze 中记录；字体来源均为固定官方 commit，保留 OFL 原许可。不包含专有系统字体。

| 本次结果 | 范围 |
| --- | --- |
| 功能回归 | 当前全套 428/428；独立 R2 定向 44/44 |
| 包消费 | 11 个 `0.1.0-issue05.20261002.2` 本地包，干净目录/缓存验证；包含本票真实字体样例 |
| 自动渲染 | full/subset、native/default 各两页 PDF；全部 Poppler exit 0 且无 stderr；full/subset 与 native/default 逐像素相等 |
| SVG | 8 个页面由真实 Chromium 渲染，FontFaceSet loaded 且 CSS 平台字体 `isCustomFont=true`；保留页面截图和字体记录 |
| PNG 目视 | 检查 full/subset PDF 1–2 页、native/default DOCX→本次 OFD→PDF 1–2 页；subset SVG Chromium 1–2 页，无缺字、裁切、重叠或样式扩散；其他 SVG 截图供复查 |
| macOS Preview | **未完成**：已观察 04 的 graphics.pdf 窗口使用中，05 正等协调独占时段；PNG/浏览器检查不替代此门禁 |

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

已保留最初样例标题过长及修正后的记录；当前两页页标已拆行，重生成后再次检查。非 BMP P2 的失败轮次和独立 R2 复验亦保留。范围限制详见[字体子集契约](../../font-subset-and-reuse.md)：CFF/变量/未知表/RTL 全量保护；非 BMP PDF 与未建模 CGTransform PDF/SVG 明确失败；已有未知 OFD 资源不裁剪。未完成 Preview 前不宣称本票整体视觉闭环或发布验收通过，也不关闭 issue #3。
