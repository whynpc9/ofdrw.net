# R01 PageBlock、文字与颜色验收

日期：2026-10-22（Asia/Shanghai）；分支 `codex/r01-pageblock-text-color`；[PR #29](https://github.com/whynpc9/ofdrw.net/pull/29)。本轮产品源码为 `e9a7390b5efc3ccdcbde6d93919325608172dcb4`，已合入主线 `62887ad6af62605411e1968da662be48ce1c091c`（R20、R06）。macOS Arm64、.NET SDK 10.0.401。复查文件保存在已忽略的 `artifacts/roadmap-v12/R01/`；[hashes.txt](hashes.txt) 仅列最终输入、输出及验收记录。

## 行为边界

- Reader 将普通页和模板页的 `PageBlock` 绘图投影为只读子元素，保留原始 XML 和文字提取；PDF、SVG、HTML 按原顺序呈现支持的路径、图片、文字、DrawParam 和有限颜色空间。未知绘图、资源歧义、非法词元或不支持的样式在输出前拒绝。
- 字体、颜色空间、DrawParam 及其引用按正整数 ID 解析；`030` 与 `30` 等值。重复字体 ID 仍在源码身份捕获阶段拒绝。PageBlock 与复合资源中的有效文字及兄弟文字可提取、可渲染；异常字体声明保留逻辑文字，但阻止三个渲染器写出。
- 普通绘图及裁切路径的弧线要求正半径和字面 `0`/`1` 标志位。裁切路径的非法 `2 2`、`1.0 0` 在 PDF、SVG、HTML 写出前统一拒绝；合法弧线的蓝色裁切区域完整呈现。空白、重复或语法不完整的路径数据也拒绝。XML 注释之间的路径空白按实际绘制文本保留。
- 颜色调色板的每个条目均经过校验，直接 `Value` 绘制不绕过资源结构检查；合法 `Value` 优先于 `Index`。隐藏元素、布尔空白、复合对象放置与裁切、局部文字样式均由针对性测试覆盖。

## 功能与包消费

- 当前源码的完整 Release 解决方案：**3560 passed、0 failed、0 skipped**，11 个测试项目；R01 PageBlock **167/167**。包含 Core `netstandard2.0` 构建和 R20 新增归档覆盖测试。使用单节点构建与仓库独占锁。
- 本地版本 `0.1.0-r01.p4` 的 **12 个基础 SDK 包和 1 个可选 SES 包**通过来源验证及独立消费者 E2E。显式 `native` 与默认 DOCX→OFD→PDF 各生成 V1/V4 测试签章，原文完整保留；签章前后逐页 PNG 一致。本地 Core、Packaging、Reader 精确包的 R18 附件清理消费者也通过，物理删除后回读为空。包未发布；默认密码学验证结果仍为 `Unsupported`。

## 页面验收

| 本轮生成链路 | 范围与结果 |
| --- | --- |
| `numeric-font-ids.ofd` → PDF → 144 DPI PNG | 1/1 页，`pageblock font` 与 `composite font` 两处文字完整、间距正常；SVG/HTML 同样含两段文字。 |
| `pageblock-clip-arc.ofd` → PDF → 144 DPI PNG | 1/1 页，蓝色半圆顶部及直边完整，无异常裁切或重影；合法路径在 SVG/HTML 中保留相同弧线。 |
| 数字绘制资源、普通弧线、XML 注释空白三个 OFD → PDF → 144 DPI PNG | 各 1/1 页，本轮 PNG 与前轮逐页 SHA-256 相同；红色方块、蓝色上半圆、蓝色三角形的前轮 Preview 结果仍可复查。 |
| `generated-layout.docx` → 本地许可字体样本 → 显式 native OFD → PDF；默认模式单独重复 | V1/V4 各两页，四组八张 96 DPI PNG 与前轮逐张 SHA-256 相同。本轮逐页 PNG 复查：中文及英文、英文斜体、蓝色表格底色和边框、第二页局部红色粗体与分页符合样例；未见缺字、裁切、重叠、重影或异常空白页。 |

macOS Preview 本轮逐页实查上述五份 PageBlock PDF 各 1/1 页，以及 `native-v1.pdf`、`default-v4.pdf` 各 1–2 页，共七份 PDF、九页。五个图元和字体样例清晰完整；DOCX 的中文、英文斜体、蓝色表格底色与边框、第二页红色粗体及日期符合样例，未见缺字、裁切、重叠、重影或异常空白页。查看链路、PDF 与逐页 PNG 哈希、源码和当前 head 的仅文档差异，记于 `artifacts/roadmap-v12/R01/p4-preview-release.json` 及中央 `gui-release-R01-20261022T0128Z.json`；已关闭全部窗口并退出 Preview。结论仅覆盖上述样例和实查页面；外部 OFD 阅读器、V5/GBT 端到端互操作、签名法律效力及任意复杂 DOCX 的保真性未验收。
