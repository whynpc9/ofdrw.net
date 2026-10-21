# R01 PageBlock、文字与颜色验收

日期：2026-10-22（Asia/Shanghai）；分支 `codex/r01-pageblock-text-color`，[PR #29](https://github.com/whynpc9/ofdrw.net/pull/29)。产品源码提交 `51772fcab892374e58361101ec4cf9921bd4ad63`，已合入主线 `fc5c96afbc41442cfa59c049400634ba08e1b8a9`（R16）。测试环境为 macOS Arm64、.NET SDK 10.0.401。输入、日志和输出在已忽略的 `artifacts/roadmap-v12/R01/`；[hashes.txt](hashes.txt) 记录本轮必要文件的 SHA-256。

## 范围与行为

- Reader 对普通页和模板中的 `PageBlock` 建立只读子元素投影；保存仍以原始 XML 为准，Split 后维持投影与文字提取。PDF、SVG、HTML 对未解析绘图、外来资源命名空间、非法原始标量及不支持的绘制参数明确拒绝。
- 有界颜色解析覆盖 DrawParam 相对继承、8-bit GRAY/RGB、无 ICC 的简单 CMYK 与调色板；文字填充、描边、横向缩放和支持的字体轮廓分别呈现。缺失或歧义引用、渐变/图案及不支持的文字语义继续拒绝。
- `PageBlock` 内的 `CompositeObject` 使用**原始**放置 XML 验证正宽高 `Boundary` 与包括 `CTM` 的标量词元。无效放置仍投影资源文字及同页、后续页文字，但 PDF/SVG/HTML 的预检均拒绝渲染；有效放置正常呈现。四种测试分别覆盖有效、缺失边界、零宽边界与非法矩阵。

## 功能测试与包消费

- 合并 R16 后的完整 Release 解决方案 11 个测试项目 **3388 passed、0 skipped、0 failed**；R01 PageBlock **128/128**；复合对象提取与渲染回归 **4/4**。Core `netstandard2.0` 包含在完整构建内。单节点构建使用仓库规定的 CLI 环境与共享独占锁。
- 本地版本 `0.1.0-r01.r16merge4` 的 **12 个基础 SDK 包和 1 个可选 SES 包**通过来源验证及独立消费者 E2E；显式 native 和默认 DOCX→OFD 各生成 V1/V4 测试签章，再从各自 OFD 导出 PDF。原始文本完整，四份签章 PDF 与签章前基线逐页 PNG 一致。默认密码学验证为 `Unsupported`；本地包未发布。

## 自动渲染、PNG 与 Preview 视觉验收

| 本轮产物与实际查看链路 | 实查范围 | 结果 |
| --- | --- | --- |
| `pageblock-composite-valid.ofd` → PDF → 144 DPI PNG → macOS Preview | 1/1 页 | 黑色方块位置与大小符合样例；无裁切、重影或空白页 |
| 普通/模板 PageBlock、图片、三种合法媒体资源与相对 DrawParam OFD → PDF → 144 DPI PNG | 七个样例各 1/1 页，逐页 PNG | 蓝色路径/文字、红色图片、黄底蓝框及局部洋红文字可见；普通/模板和三种媒体页面各自一致；与 R16 合并前本轮对应 PNG 哈希一致 |
| `generated-layout.docx` → 本地许可字体替换样本 → 显式 native OFD → PDF → Preview；默认模式单独重复 | 本轮 `native-v1.pdf` 和 `default-v4.pdf` 各 1–2 页；四种 V1/V4 模式各两页 PNG | 中英文、英文斜体、蓝色表格底色/边框、第二页局部红色粗体与分页符合样例；未见缺字、裁切、重叠、重影或异常空白页；八张 PNG 按页哈希相同 |

Preview 查看的是本次生成的 PDF，均来自相应 OFD；三份文档窗口已关闭，Preview 进程已退出。检查路径、页码、哈希与源码身份记录在 `gui-release-R01-20261021T1752Z.json`。`rendered-pages.json` 的 `previewAccepted: false` 是自动化阶段的状态，独立 Preview 验收以该 GUI 记录为准。签章前后 PNG 等同性属于自动渲染证据，不代表逐份签章 PDF 均已在 Preview 查看。`native-v1.ofd` 为 11,605,702 字节，对应 PDF 为 118,950 字节，未见明显体积变化。

结论仅覆盖上述样本与实际检查的页面。原始上游 `999.ofd` 的无效章内颜色空间引用仍明确拒绝；另制签名失效呈现副本的先前结果不作为原件互操作证明。外部 OFD 阅读器、V5/GBT 端到端互操作、签名法律效力及任意复杂 DOCX 的保真性均未验收。
