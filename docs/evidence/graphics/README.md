# 04 原生绘图验收证据

源码冻结于 `055f43affa6e288ecd45ac446a11d19d68d59123`，基线为 PR #12 的 `df71c7f20e0c45e9cba9cb78f2d2a91d61413ec9`。最终审核状态见 [acceptance.json](acceptance.json)，文件尺寸/SHA-256/环境及模式见 [manifest.json](manifest.json)。本票仍等待 Preview 独占时段和最新 head 外部评审/CI，不宣称整体通过。

| 检查层 | 已完成范围 | 结果 |
| --- | --- | --- |
| 功能回归 | 全套 410 项，独立新增 10 项；Native 对象往返/原文/矩阵/状态/裁剪/字体身份/异常预算取消 | 通过 |
| API 文档 | 新 Graphics 公开 API CS1591 检查 | 0 新警告 |
| 本地包消费 | 11 个临时本地包实际安装/转换；未发布 | 通过 |
| 自动渲染 | 四组各两页：graphics、graphics-roundtrip、baseline-native、baseline-default；PDF 8 页、SVG 8 页 | 全部生成并验证 |
| PNG 视觉 | 本次 16 页 PDF/SVG PNG；中英、比例字距、局部粗斜/红色、框线底色、基线、曲线、非等比线宽、填充洞、clip、旋转及分页 | 检查范围内通过 |
| macOS Preview | 必须是本次原生 OFD → PDF，包含显式 Native/default 基准；独占时段尚未确认 | **未完成**，PNG 不替代此门 |
| 最新 head Codex/Cursor/CI | stacked PR 的实时结果 | 待外部审核 |

图形样例仅使用 Layout（传递 Core/Packaging），实际包内 17 个 PathObject、29 个 TextObject、1 份字体、0 个 ImageObject。两页为中英票面和变换示意图，不是整页位图。Native/default 各两页使用固定 `generated-layout.docx` 的许可字体变体，OFD 原文完整；查看链路是 DOCX → Native/default OFD → PDF，未以直接 DOCX→PDF 代替。

样例字体是仓库固定 installer 的 Noto Sans CJK SC Regular，随归档附 OFL 许可。不包含专有系统字体。04 未子集化，单 OFD 约 11.6 MB，大小由完整嵌入字体主导；roundtrip 增量仅十余字节，Native/default 同尺寸。05 负责真实子集与资源复用；此票不宣称完成 issue #3 全部要求。

`evidence.tar.zst` 包含本次 OFD/PDF/SVG、16 个逐页 PNG、原文、源 DOCX、OFL 字体/许可、日志及清单。可用 `tar --zstd -xf docs/evidence/graphics/evidence.tar.zst -C artifacts/graphics-review` 提取，之后按 manifest 校验文件大小和 SHA-256。若系统 tar 不支持 --zstd，使用 `zstd -d -c ... | tar -xf -`。

复现：安装固定许可字体到 `artifacts/graphics-fonts`，执行 `scripts/run-graphics-e2e.sh artifacts/graphics/current`；.NET 采用用户 AGENTS 要求的可写 CLI_HOME、环境变量和单节点 flags。PDF PNG 由 Poppler 110 DPI 生成，SVG 由 librsvg 渲染在白色背景。上下文、单基线文字和不支持操作边界见 [设计契约](../../graphics-design-contract.md) 与 [教程](../../tutorials/16-native-graphics.md)。只对实际生成/检查样例作结论。
