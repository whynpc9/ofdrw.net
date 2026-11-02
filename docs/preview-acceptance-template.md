# Native OFD → PDF → macOS Preview 验收记录模板

> 复制到 `artifacts/<candidate>/preview-acceptance.md` 并填写真值。空白、`未检查`、自动化绿测均不表示 Preview 已验收。每次重新生成 OFD/PDF 后，关闭并重新打开 Preview 文档，再逐页复查。

## 受限 profile 与独立限制

preview.10 采用用户 2026-11-02 选择的 `cn-en-physical-cjk-400-700-v1`；详见[当前计划](next-steps-2026-11-02.md)及 [Q05 说明](release-engineering-gates.md)。本模板不预填通过结果。

| 项目 | 实际值 |
| --- | --- |
| release profile ID、承诺范围 | 未填写 |
| 显式 FontDirectories、源字体 manifest/三字体哈希与实际选中 face/子集证据 | 未填写 |
| 严格 Throw、Office 开关、固定 fallback 与两模式实际参数 | 未填写 |
| 旧 Regular-only 对照/量字 FAIL 的独立证据与边界 | 未填写 |
| Preview Find 的实际 OS/build、UI 观察、已知兼容性限制 | 未填写 |
| 原始 Unicode Copy 的实际验证状态/工具限制 | 未填写 |
| 用户 2026-11-02 范围决定与 known limitations 披露位置 | 未填写 |

最终 JSON 保留全部核心字段，并另附 `release_profile` 对象和 `known_limitations` 列表。`open_findings=[]` 只可表示该明确 profile 的本次必验页面无开放视觉缺陷；不能宣称范围外旧 FAIL、Preview Find 或未完成原始 Copy 已通过。脚本不验证附加字段语义，须由协调者/Astra 对原证据复核。禁止提前改旧 accepted JSON，禁止将新的 profile 内失败改列限制来放行。

## 候选与产物

| 项目 | 实际值 |
| --- | --- |
| 源码提交 SHA、分支、包版本（与发布 tag 一致） | 未填写 |
| 测试样例及来源、SHA256、隐私状态 | 未填写 |
| 转换命令与显式模式（`native`/默认） | 未填写 |
| 本次显式 Native OFD 路径、SHA256、字节数 | 未填写 |
| 由该 Native OFD 导出的 PDF 路径、SHA256、字节数 | 未填写 |
| 本次默认模式 OFD 路径、SHA256、字节数 | 未填写 |
| 由该默认模式 OFD 导出的 PDF 路径、SHA256、字节数 | 未填写 |
| 页面 PNG 路径、SHA256、渲染工具与 DPI | 未填写 |
| 先前可比产物体积及变化 | 未填写 |
| 功能测试命令、结果、日志位置 | 未填写 |

## 实际查看链路

- [ ] `DOCX → 显式 native OFD → PDF → macOS Preview`。不得以直接 DOCX→PDF 替代。
- [ ] `DOCX → 默认 OFD → PDF → macOS Preview`，与显式 Native 分别记录产物。
- [ ] 检查人、日期时间、macOS/Preview 版本：未填写。
- [ ] 重新载入/关闭后重新打开本次 PDF，排除 Preview 旧页面缓存：未检查。
- [ ] OFD 原始文本完整；PDF 导出没有重复文本：未检查。

## 逐页检查

| 模式 | 页码/关键区域 | 中文缺字/乱码 | 英文间距/断词 | 粗斜体/颜色/字号/基线 | 表格填充/边框/对齐 | 分页/裁切/重叠/重影/空白页 | 图片/页眉页脚/页码 | 结果与截图 |
| --- | --- | --- | --- | --- | --- | --- | --- |
| 未检查 | 未检查 | 未检查 | 未检查 | 未检查 | 未检查 | 未检查 | 未检查 | 未检查 |

## 结论

- 已检查页面：显式 Native **未检查**；默认模式 **未检查**。
- 功能测试：**未执行**。
- 像素差门禁：**未执行**；语料、逐页 RMSE、阈值和报告路径：未填写。
- Preview 视觉验收：**未完成**。
- 遗留缺陷、未检查页面和原因：未填写。
- 回滚依据（上个已验证标签/包）、候选放行人：未填写。

只有上述链路中的本次页面被 Preview 逐页查看，且明确 release profile 内视觉缺陷关闭后，才能把该 profile 的视觉验收记为完成；Find/Copy 限制及范围外旧 FAIL 继续单独披露。CI 可以验证产物和像素差，无法代替这项人工记录。
