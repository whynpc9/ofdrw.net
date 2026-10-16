# PR23 两项修复后 PNG 与 macOS Preview 验收

本次冻结源码 `979a93e39dd8cec3fbb4bdd2121d301ea84b2e39`；清单 SHA256 `6da5cc063ac6b67fbf3fb99fa5aa717dc248719bb6dd67a3b22447f9070d7976`。8 份 PDF / 14 页已逐页实际查看 PNG 和 macOS Preview，限定本清单样例 **PASS**。

## 查看路径与范围

| PDF | 已查看页码 | PNG / Preview |
| --- | --- | --- |
| [generated-docx-native.pdf](/Users/wanghongyi/.codex/worktrees/audit-followup-coordination/ofdrw.net/artifacts/pr23-review-fixes/package-output/generated-docx-native.pdf) | 1, 2 | PASS / PASS |
| [generated-docx-default.pdf](/Users/wanghongyi/.codex/worktrees/audit-followup-coordination/ofdrw.net/artifacts/pr23-review-fixes/package-output/generated-docx-default.pdf) | 1, 2 | PASS / PASS |
| [cli-native.pdf](/Users/wanghongyi/.codex/worktrees/audit-followup-coordination/ofdrw.net/artifacts/pr23-review-fixes/package-output/cli-native.pdf) | 1, 2 | PASS / PASS |
| [direct-built-in.pdf](/Users/wanghongyi/.codex/worktrees/audit-followup-coordination/ofdrw.net/artifacts/pr23-review-fixes/visual/direct-built-in.pdf) | 1, 2 | PASS / PASS |
| [selected.pdf](/Users/wanghongyi/.codex/worktrees/audit-followup-coordination/ofdrw.net/artifacts/pr23-review-fixes/visual/selected.pdf) | 1 | PASS / PASS |
| [notice.pdf](/Users/wanghongyi/.codex/worktrees/audit-followup-coordination/ofdrw.net/artifacts/pr23-review-fixes/visual/notice.pdf) | 1, 2 | PASS / PASS |
| [italic-boundaries.pdf](/Users/wanghongyi/.codex/worktrees/audit-followup-coordination/ofdrw.net/artifacts/pr23-review-fixes/visual/italic-boundaries.pdf) | 1 | PASS / PASS |
| [external-libreoffice.pdf](/Users/wanghongyi/.codex/worktrees/audit-followup-coordination/ofdrw.net/artifacts/pr23-review-fixes/visual/external-libreoffice.pdf) | 1, 2 | PASS / PASS |

显式 native、默认和 CLI：`generated-layout.docx → native OFD → PDF → macOS Preview`。直接 BuiltIn 与本次真实 external LibreOffice 为各自 `DOCX → PDF → Preview` 路径，单列检查；notice 为纯文本 → native OFD → PDF；selected 为 native 选源第 2 页；italic-boundaries 为 BuiltIn 真实 italic/fallback 边界样例。没有以直接 DOCX→PDF 代替 native 产物。

## 页面观察

native/default/CLI 各两页：中英文完整可见；蓝标题、灰色 italic 副标题、红色局部强调、表格底色和网格、右对齐日期、正文基线与分页正常。selected 只有源第二页，未带入额外空白页。

直接 BuiltIn 两页：副标题斜体有边界，常规正文直立，表格底色和黑边框可见；其中文标题及红色文字字重比 native/LibreOffice 较轻，保留该引擎现有布局/字重差异，不宣称与其他引擎逐像素或严格字重一致。真实 external LibreOffice 两页：中英文、标题/斜体副标题、蓝表格边框与填充、第二页粗红局部强调和日期完整可见。

notice 两页：比例英文长行换行和空行正常，第一页01–04，第二页05–24和结束语完整。Preview 明确看到23、24各自独立的行和 End of notice，无重叠或裁切。italic-boundaries：相同文字 regular / 红 italic / regular 边界清楚；真实 Arial italic 无明显二次倾斜，fallback 模拟 italic 可见；混排换行后恢复 regular，中文及英文间距完整。

未发现本次范围新增缺字/乱码、样式扩散、重叠、重影、异常空白或裁切。该 PDF 清单没有图片、页眉页脚和页码专项样例，未据此扩展结论；HTML mixed 另有图片/clip检查。

## 功能与自动渲染

本代理只做只读验收和证据保存，未运行构建/测试/包消费、未改源码、未重生成 PDF/HTML。父 root 单独负责功能门；本清单 automatic_render=PASS，14张PNG与PDF当前哈希逐一匹配。自动化、文本可提取及数量均没有代替本次目视和Preview。

## 原始证据

时间：2026-10-16T06:00:19.342818+00:00。`pages.json` 保存14页完整file URL、实际Page AX、原始AX文件与JPEG名/大小/SHA及观察。所有JPEG由CUA原生Preview截图字节直接落盘，验证JPEG文件头，未重压缩或改扩展名。`sample-expectations.json`保留DOCX run属性/样例身份；PNG独立记录在`../png-review.json`。

`pdf-before.json` / `pdf-after.json`确认8份PDF前后SHA完全一致；清单本身未修改。`cleanup.json`逐个记录关闭前核对的8个当前任务URL；最后CUA返回noWindowsAvailable。初始为Open面板，未关闭任何用户已有文档；没有使用voice-only capture_screen。GUI租约贯穿本次Preview和HTML验收，最后由本持有者释放。

本结论绑定该冻结源码和产物SHA；若再改生产源码并重新生成文件，需要新产物验收。
