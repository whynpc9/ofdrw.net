# R01 PageBlock、文字与颜色验收

日期：2026-10-22（Asia/Shanghai）；分支 `codex/r01-pageblock-text-color`，[PR #29](https://github.com/whynpc9/ofdrw.net/pull/29)。本轮源码为 `00bba7d14f75efff55fd8be0e539fd6f2c0f54b2` 加本 PR 的复合对象改动；这两份实现/测试文件相对该 SHA 的 diff SHA-256 为 `d4826292c83ee9f911ee1940144cf678bf94c29e74e5f16cf219d7fbb84ff5a7`。合并前最终提交身份以 PR head 为准。环境：macOS Arm64、.NET SDK 10.0.401。产物与日志保存在已忽略的 `artifacts/roadmap-v12/R01/`；[hashes.txt](hashes.txt) 只列本轮输入、输出和支撑结论的记录。

## 范围与行为

- Reader 对普通页和模板中的 `PageBlock` 建立只读子元素投影，原始 XML 仍是保存来源；Split 后保持该投影及文字提取。PDF、SVG、HTML 对未解析绘图、外来资源命名空间、非法原始标量及不支持的绘制参数明确拒绝。
- 有界颜色解析覆盖 DrawParam 相对继承、8-bit GRAY/RGB、无 ICC 的简单 CMYK 与调色板；文字填充、描边、横向缩放和支持的字体轮廓分别呈现。缺失或歧义引用、渐变/图案及不支持的文字语义继续拒绝。
- 本轮修复：在展开 `PageBlock` 内的 `CompositeObject` 资源前，验证**原始**放置 XML 中存在有限正宽高的 `Boundary`，并验证包括 `CTM` 在内的已知标量词元；Reader 的默认值不得把非法放置变成可绘制对象。有效复合对象仍投影。测试覆盖有效放置、缺失边界、零宽边界和非法矩阵；拒绝路径在 PDF/SVG/HTML 均抛出 `NotSupportedException`。

## 功能与包消费

- 定向复合对象回归 **4/4**；全部 R01 PageBlock 回归及成功产物生成 **128/128**；完整 Release 解决方案 11 个测试项目 **3350 passed、0 skipped、0 failed**。Core 的 `netstandard2.0` 包含在完整构建内。单节点构建使用仓库规定的 CLI 环境与共享独占锁。
- 本地版本 `0.1.0-r01.composite2` 的 **12 个基础 SDK 包和 1 个可选 SES 包**完成来源验证与独立消费者 E2E；显式 native、默认 DOCX→OFD 各自生成 V1/V4 测试签章，再从各自 OFD 导出 PDF。原始文本完整，四份签章 PDF 与各自签章前基线逐页 PNG 一致。默认密码学验证为 `Unsupported`；本地包未发布。
- 首次本轮包消费因默认样本字体目录不存在而在启动时停止；使用仓库已有的字体与许可文件目录，在全新输出目录重跑通过。该前置资料问题未计为产品回归。

## 页面视觉验收

| 本轮产物与查看链路 | 实查范围 | 结果 |
| --- | --- | --- |
| `pageblock-composite-valid.ofd` → PDF → 144 DPI PNG → macOS Preview | 1/1 页 | 黑色方块位置与大小符合有效放置样例；无裁切、重影或空白页 |
| 普通/模板 PageBlock、图片、三种合法媒体资源与相对 DrawParam OFD → PDF → 144 DPI PNG | 各 1/1 页，逐页 PNG；其中图片、媒体及颜色页重新查看 | 蓝色路径/文字、红色图片、黄底蓝框及局部洋红文字可见；普通/模板图像相同，三种媒体图像相同；与上一轮 Preview 已检查的对应页面 PNG 哈希一致 |
| `generated-layout.docx` → 本地许可字体替换样本 → 显式 native OFD → PDF → Preview；默认模式单独重复 | 本轮 `native-v1.pdf` 和 `default-v4.pdf` 各 1–2 页；四种 V1/V4 模式各两页 PNG | 中英文、英文斜体、蓝色表格底色/边框、第二页局部红色粗体与分页符合样例；未见缺字、裁切、重叠、重影或异常空白页；八张本轮 PNG 逐页哈希按页相同 |

本轮 Preview 文档窗口均已关闭，Preview 进程已退出；检查路径、页码和哈希另记在 GUI 释放记录 `gui-release-R01-20261021T1722Z.json`。签章前后 PNG 等同性是自动渲染证据，未冒充逐份 Preview 查看。`native-v1.ofd` 为 11,605,730 字节（上一轮 11,605,713），对应 PDF 均为 118,950 字节，无明显体积变化。

结论仅覆盖上述样本与实际检查的页面。原始上游 `999.ofd` 的无效章内颜色空间引用仍明确拒绝；另制签名失效呈现副本的先前结果不作为原件互操作证明。外部 OFD 阅读器、V5/GBT 端到端互操作、签名法律效力及任意复杂 DOCX 的保真性均未验收。
