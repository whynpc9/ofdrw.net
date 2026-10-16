# 整合候选 PNG / macOS Preview 验收

结论：本次清单 **7 份 PDF / 12 页全部通过** PNG 目视与原生 Preview 验收。结论仅覆盖下列实际查看页面。

- 绑定源码：`aef66e76d923d706bcd440eb1583c906f300eeb7`，协调分支 `codex/audit-integration`。
- 清单：`artifacts/followup-integration/visual-manifest.json`；本次快照 SHA256：`2cfc5a3e148333d8599f0b9d8207accc62e6e251e69b3877b969cf4d93cf7adb`。
- 实际产物均位于本协调 worktree；每次打开后核对完整 file URL，未复用原003或其他旧会话的已验 PDF。
- 本轮只做读取、视觉验收与证据保存；未修改产品源码、未重新生成 PDF、未运行 build/tests/packages、未提交文件。
- 原003分支 HEAD保持 `0fda098f43d6662cfdf2fde1a4ba5b2e3f00df8b`；原003.json未修改。
- 时间：2026-10-16T11:26:09.966591+08:00；macOS Preview实际页面检查，原始JPEG保存，无重压缩/缩放。

## 已检查范围

| 实际PDF | 已看页码 | PNG / Preview |
| --- | --- | --- |
| [generated-docx-native.pdf](/Users/wanghongyi/.codex/worktrees/audit-followup-coordination/ofdrw.net/artifacts/followup-integration/package-output/generated-docx-native.pdf) | 1, 2 | PASS / PASS |
| [generated-docx-default.pdf](/Users/wanghongyi/.codex/worktrees/audit-followup-coordination/ofdrw.net/artifacts/followup-integration/package-output/generated-docx-default.pdf) | 1, 2 | PASS / PASS |
| [cli-native.pdf](/Users/wanghongyi/.codex/worktrees/audit-followup-coordination/ofdrw.net/artifacts/followup-integration/package-output/cli-native.pdf) | 1, 2 | PASS / PASS |
| [direct-built-in.pdf](/Users/wanghongyi/.codex/worktrees/audit-followup-coordination/ofdrw.net/artifacts/followup-integration/visual/direct-built-in.pdf) | 1, 2 | PASS / PASS |
| [selected.pdf](/Users/wanghongyi/.codex/worktrees/audit-followup-coordination/ofdrw.net/artifacts/followup-integration/visual/selected.pdf) | 1 | PASS / PASS |
| [notice.pdf](/Users/wanghongyi/.codex/worktrees/audit-followup-coordination/ofdrw.net/artifacts/followup-integration/visual/notice.pdf) | 1, 2 | PASS / PASS |
| [italic-boundaries.pdf](/Users/wanghongyi/.codex/worktrees/audit-followup-coordination/ofdrw.net/artifacts/followup-integration/visual/italic-boundaries.pdf) | 1 | PASS / PASS |

实际链路：显式 native / 默认 / CLI 的 DOCX → native OFD → PDF → Preview；直接 BuiltIn DOCX → PDF 单列；纯文本 → native OFD → notice.pdf；真实italic边界PDF单列。

## 页面观察

显式native、默认和CLI各两页：中英文完整可读；副标题斜体、蓝色标题、红色强调文本、表格填充/边框/对齐、日期和分页正常。直接BuiltIn两页：副标题斜体保留，邻接regular正文直立，无样式扩散；保持该引擎既有布局差异。选页PDF仅含源第二页，内容和日期正常。

真实italic边界页：同文不同样式的相邻regular与italic分界清楚，真实italic未出现明显二次倾斜，regular fallback的模拟italic可见；换行及中文完整，未见裁切、重叠或重影。

notice两页：长行换行、空行、比例英文字距、中英文基线及分页正常；01-04在第一页，05-24与结束语在第二页。初次PNG判读曾怀疑末页23/24行重叠；随后对同一SHA的原尺寸PNG和真正Preview页面复查，两行均完整且独立，疑点排除，未改动或重生成文件。不能把文本可提取当视觉验收，本次已实际查看两行及结束语。

该清单未另含图片、页眉页脚或页码控制，未据此推断这些场景或任意复杂Word文档的保真程度。

## 原始证据

- `pages.json`：12页真实完整URL、页码AX摘要、原始AX JSON、JPEG文件名/字节数/SHA256和逐页观察。
- 每页 `.ax.json` 保留未改写的完整AX文本；`.jpg`为原始JPEG捕获字节。
- `pdf-before.json` / `pdf-after.json`：7份PDF验收前后SHA256，全部一致；PNG也逐项核对当前清单SHA。
- `after-close-current.ax.json` / `after-close-window-menu.ax.json`：仅关闭本轮打开的7份PDF标签，保留用户已有窗口/标签。
- PNG逐页记录在上一级 `png-review.json`。4份HTML未由本会话浏览，交029后续完成。

## 保留问题

既有legacy源landscape页面设置的二次交换问题仍未扩修；本轮清单不证明该问题已修复，不应据此宣称任意横向源文档均通过。

本轮PDF标签已全部关闭，GUI独占租约已释放；029可取得租约继续HTML浏览器验收。
