# 本地验收结果 · 2026-09-06

验收候选：`0.1.0-preview.fixes.20260906.3`，基于 `3d68d67fd382044df7c17201980573d989384fba` 及本次提交的修复。验证期间生产源码未变化；提交前回读候选源码与 50 个视觉产物哈希一致。

- Release 构建：0 warning / 0 error。
- .NET 回归：82/82 通过，0 失败、0 跳过。Core 5、Packaging 22、PDF/SVG 20、Signatures 4、DOCX 26、CLI 5。
- 包验证器：5/5 通过；11 个包在私有缓存中实际安装消费，包含 CLI；消费后清单复验通过。
- macOS Preview：10 份最终 PDF、23 页逐页查看通过。Native 验收采用 DOCX → native OFD → PDF → Preview；未用直接 DOCX→PDF 替代。

| 样例 | 已查看页 | 结论 |
| --- | --- | --- |
| generated-docx-native / generated-docx-default | 各 1–2 | 中英文、比例字体、局部粗斜体/颜色、表格、右对齐日期和分页正常 |
| headers-native | 1–5 | 首/奇/偶页眉，正文与页脚分离，分节页码 1、2、3、10、11 正确 |
| image-native | 1 | 红蓝 JPEG 比例、位置正常 |
| merged-styles | 1–2 | 同名字体的人工矩形/三角形分别保留；图片旋转、半透明与裁剪正确 |
| supplemental-native / supplemental-default | 各 1 | 中英文脚注、尾注、批注原文保留于标签下 |
| selected-native | 1–3 | 源页选择 [2,0,2]（零基）显示 three、one、three |
| automatic-dual-layer | 1–3 | ROW-00…44 按实际分页分布，透明语义文字不露出 |
| selected-dual-layer | 1–3 | 选择 [1,0,1] 后文字与图像对应，无重影 |

检查范围内未发现残留缺字、裁切、重叠或异常空白页。附属文字标签追加不等于 Word 页底脚注排版；结论不推广到任意复杂 Word 文档。详细能力边界见 [转换契约](conversion-contracts.md)。

本机基准（ARM64、.NET 10.0.7、150 DPI、预热后 3 次中位数）：100 页 PDF 耗时 2128.18 → 1868.97 ms（-12.2%），托管分配 836.61 → 537.06 MiB（-35.8%）；111 个 PDF PNG 载荷相同。10 页耗时 +1.3%；DOCX 耗时 -11.9%、托管分配 +4.2%。峰值工作集返回 0，无法据此评估原生内存峰值。100 页 OFD 体积约 +0.21%；DOCX A4 栅格高度由 1754 变为 1753 像素，因此不声称该样例像素完全一致。

完整本地产物和日志保存在 `artifacts/repository-fixes-2026-09-05/`（Git 忽略）：`acceptance.md`、`final-tests3.log`、`final-tests3/trx/`、`package-final3.log`、`visual-final/`、哈希清单和 `benchmark/`。这里记录的是本地执行结果，远端 CI 需以推送后的运行结果为准。
