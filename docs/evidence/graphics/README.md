# 04 原生绘图验收证据

R4 源码冻结于 `7b156032eb7ab2dd78f0f0ba33b0359dd4eea3a1`，基线为 PR #12 的 `df71c7f20e0c45e9cba9cb78f2d2a91d61413ec9`。[PR #13](https://github.com/whynpc9/ofdrw.net/pull/13) stacked base 为 `codex/document-tools`。状态见 [acceptance.json](acceptance.json)，尺寸/SHA-256/环境/模式见 [manifest.json](manifest.json)。本次 R4 产物已重新生成；**R4 Preview 与最新 head 复审待完成**，不以 R1 通过替代。

| 检查层 | R4 本次范围 | 结果 |
| --- | --- | --- |
| 功能回归 | 全套 425 项；独立新增合计 25 项原生往返/原文/矩阵/状态/clip/字体绑定/预算取消、数字/强调及栅格回归 | 通过 |
| API 文档 | Graphics 新公开 API CS1591 | 0 新警告 |
| 本地包消费 | 11 个 `0.1.0-graphics.20261003.r4final` 包真实重建、安装和转换 | 通过；未发布 |
| 自动渲染 | graphics/roundtrip/Native/default 各两页，加 name-only/分数clip各一页；PDF 10 页/SVG 10 页 | 全部生成验证；尖角实际像素及默认 limit 负对照通过 |
| PNG 视觉 | 20 页 PDF/SVG，全页复查 | 检查范围内通过 |
| macOS Preview | R4 本次 OFD→PDF 10 页；已申请下一独占时段 | **未完成**；PNG 与 R1 不替代 |
| 最新 head 外部审核 | 首轮5条及R2新增2条意见统一修复，需新 head 复审与 CI | 待闭环 |

R4 修复可写矩阵在三位小数下退化，路径/clip/DeltaX 的科学计数法，SVG 已声明 cap/join/miter-limit，以及 name-only 斜体遇已有用户 CTM 时未输出原生 shear。Graphics 游程归一基线；Writer 仅对 fresh、normalized、无 DeltaY 的 name-only 文字组合 M*F，只标注生成 F，保留 03 去因子契约与 05 资源绑定，嵌入/透明/已保留 SourceXml 不重新组合。新建内部普通十进制 formatter 共用，无新公共字体服务。

公开两页样例只引用 Layout（传递 Core/Packaging），包内 18 个 PathObject、30 个 TextObject、1 份字体、0 个 ImageObject。新增 name-only OFD 只有 7 个 TextObject/4 个 PathObject 和不含 FontFile 的字体声明，原生 CTM 带明确 faux 因子。其 PDF 显式用同一固定 OFL Noto 字节注册现有 resolver；SVG PNG 使用隔离 fontconfig 目录。所有公开字体载荷仅为固定 installer 的 Noto Sans CJK SC Regular，并附 OFL；不包含专有系统字体。

显式 Native/default 基准使用本次许可字体变体 `generated-layout.docx`，OFD 原文完整。链路为 DOCX → Native/default OFD → PDF/SVG，没有使用直接 DOCX→PDF 替代。原生图形完整嵌入字体 OFD 约 11.6 MB，roundtrip 仅十余字节差；name-only 约2 KB并依赖阅读器字体环境。05 负责真实字体子集/复用；04 不宣称整个 issue #3 完成，未新增 flow/table/Canvas/Skia 适配。

`evidence.tar.zst` 含本次原始 OFD/PDF/SVG、20 页 PNG、Unicode 原文、许可 DOCX、字体/OFL、环境/哈希与日志。默认 zstd 解码兼容的 128 MiB 窗口。使用 `zstd -d -c docs/evidence/graphics/evidence.tar.zst | tar -xf - -C artifacts/graphics-review` 全新提取，再按 manifest 检查尺寸/SHA-256。R1 原始47文件归档与8页Preview记录保留于 Git commit `4b634d1`。

复现：固定 installer 安装字体到 `artifacts/graphics-fonts`，执行 `scripts/run-graphics-e2e.sh artifacts/graphics/review4-final`。Poppler PDF 110 DPI；librsvg SVG 白背景、720px宽，字体仅来自许可目录。严格采用 AGENTS 的 CLI_HOME/skip/telemetry/单节点 flags。只对列出的实际样例作结论；复杂 shaping、任意 Word 完整保真不在本票范围。

R4 使用有界 BigInteger 精确比较实际十进制系数的det，统一模型负责的Text/Path/Image/default-image与clip CTM精度，SVG矩阵/translation/glyph原点同样保真。保留Text SourceXml所有权、原有Box/Size/LineWidth三位精度。包消费者直接从NuGet feed调用Graphics并检查原生对象、0.9996CTM及faux提示。窄页PDF及普通148mm页上的100000mm逻辑框展示旧舍入丢边框；PDF/SVG实际像素断言通过，原失败SVG探针日志保留。R1、R2完整历史证据在各自Git提交中，当前归档全部63文件已全新提取并校验尺寸/SHA256。

R4 共用现有OfdTextEmphasis纯faux组合，Graphics在Commit前预检，使double.MaxValue的极端name-only italic组合当场原子失败；Writer复用相同数值计算，显式/隐式资源、透明/嵌入及预算恢复独立测试通过。正常样例重新生成全部20页PNG，与R3逐字节相同；本次10页PDF仍需Preview独占补验。历史R2计数为417，R3为424，R4为425，均独立保留。
