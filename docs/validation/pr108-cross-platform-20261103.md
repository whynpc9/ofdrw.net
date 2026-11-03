# PR #108 跨平台验证

日期：2026-11-03（Asia/Shanghai）。验证范围：Regular-only Han TextStroke 的 OFD→PDF／Image 横向强调修复。

结论：**本轮固定样例跨平台验证通过**。功能测试、自动渲染检查、PNG 视觉复查、macOS Preview 验收四项均通过。

## 源码与证据来源

- [PR #108](https://github.com/whynpc9/ofdrw.net/pull/108) head：`846dc639d9f4a698ebf79dc32f3910196611673e`。
- 本机验证合并后的 main：`8b38024152bf207717a0fc15cf3a0316626fe011`；与 PR head 的完整 Git tree 相同（`855116e60b09de8ad194b36a4a4dec15924fff78`）。
- 三平台执行来源：[CI run 41565844605](https://github.com/whynpc9/ofdrw.net/actions/runs/41565844605)。本轮下载并解析实际 TRX、包清单和页面，未重新触发远程 CI。
- 本机：macOS 26.6.2 arm64、.NET SDK 10.0.401；构建采用隔离 DOTNET_CLI_HOME、单节点、禁用 build server/shared compilation。
- 完整运行证据目录：ignored `artifacts/pr108-cross-platform/`；必要的输出、页面与 Preview 截图保存在[可选 LFS 归档](pr108-cross-platform-evidence/pr108-20261103-acceptance.zip)，大小和逐文件 SHA-256 见[清单](pr108-cross-platform-evidence/manifest.json)。原 DOCX／源字体复用仓库 fixture，不重复归档。

## 功能测试

三平台全套各发现 4659 项；下表的强制字体测试是单独执行的 74 项，不计入全套计数。

| 平台 | 全套通过 | 失败 | 跳过 | 强制字体测试 | DenseCjk + BoldAlignment |
| --- | ---: | ---: | ---: | --- | --- |
| Linux | 4641 | 0 | 18 | 74/74 | 27/27 |
| Windows | 4636 | 0 | 23 | 74/74 | 27/27 |
| macOS | 4642 | 0 | 17 | 74/74 | 27/27 |

全套中三平台均有 17 项可信字体配置条件跳过，这些测试由独立强制字体任务补验；Windows 另有 6 项 Unix 专用测试跳过，Linux 另有 1 项 Windows/macOS 专用大小写分支测试跳过。详细测试名与原因保存在 `ci-trx-summary.json`，不把跳过算作通过。

本机在合并提交重跑重点测试 27/27，0 失败、0 跳过。两个 DenseCjk 测试实际执行，覆盖字腔阈值、上横有墨、相对 Regular 墨量下限、PDF 两字无重复及覆盖层不进入横画间隙。

三平台包消费任务均成功；下载的三个 `package-manifest.json` 完全一致，12 个 `0.1.0-preview.10` 候选包 SHA-256 相同。本机另从合并源码重新构建、打包并运行完整隔离消费者：

```sh
scripts/run-converter-package-e2e.sh 0.1.0-preview.10 \
  --packages-dir "$PWD/artifacts/pr108-cross-platform/local-packages" \
  --output-dir "$PWD/artifacts/pr108-cross-platform/local-package-e2e"
```

退出码 0，`Validated 12 packages at 0.1.0-preview.10`；包括 native/dual-layer、OFD→PDF／PNG、上游样例栅格往返及比较器负对照。完整日志：`local-package-e2e.log`。

## 实际产物与自动渲染检查

复用原始 `generated-layout.docx`（SHA-256 `17bea68d57776b02e5c68c8915de11b42087809d5310f8fafcbb003db08d7bb6`），仅提供固定 Noto CJK SC Regular（`3012a9b6…`）及 Noto Sans Regular（`b85c38ec…`），关闭 Office 字体，默认 TextStroke。分别生成显式 native 和默认模式，链路均为 **DOCX → native/default OFD → PDF**。

本机复用 `Ofdrw.Net.Layout.E2E --cn-en-release-only`：两个 layout 各 2 页、headers 各 5 页、image 各 1 页、selected 各 3 页。OFD 原文与 DOCX、PDF 去空白文本逐页一致；layout 为 139+50＝189 字；样式、填充、图片原始载荷及页面选择断言均通过。选页 `[2,0,2]` 保留重复页，显示页码为 3/5、1/5、3/5；headers 节重编号为 1、2、3、10、11。

全部 8 个 PDF 均逐页渲染为 PNG。核心 layout 另通过公共 OFD→Image 144 DPI 导出，两模式都生成两页。沿用既定字腔窗口和 >100 判据，无重新调参：

| 实际查看链路 | “量”下／上字腔亮行均值（native/default 相同） |
| --- | --- |
| Linux CI PDF → Poppler，144 DPI | 200.6 / 207.2 |
| Windows CI PDF → Poppler，144 DPI | 200.6 / 207.2 |
| macOS CI PDF → Poppler，144 DPI | 200.6 / 207.2 |
| 本机 PDF → Poppler，144 DPI | 200.6 / 207.2 |
| 本机 OFD → 公共 Image（PDFium），144 DPI | 169.9 / 179.7 |

三平台 CI 内，各模式对应两页 PNG 字节一致；本机两个模式的核心及附加样例 PNG 也一致。核心 OFD 各 290726 字节、PDF 各 1649045 字节；本轮没有重新生成修复前产物，不推断修复前后体积变化。完整哈希见 `manifest.json`、`page-identities.json`，固定字体身份及转换结果见 `local-pages/cn-en-release-manifest.json`。该生成器保留的历史视觉失败说明不代表本轮验收结果，本报告和实际 GUI 记录为本轮结论。

## PNG 视觉复查

主代理逐页查看三平台 CI 的核心 native 页面 1–2（default 对应页字节一致），及本机 Poppler／PDFium 核心两页。另查看 headers 5 页、image 1 页及 selected 3 页的联系表，核对页眉页脚、图片和页码。未见中文缺字、英文断裂、局部样式扩散、裁切、重叠或异常空白页；表格底色与边框正确，“量”字腔打开，红色强调仍明显重于正文。

## macOS Preview

由当前主机模型清单中的最新正式 Terra `gpt-6.1-terra`、High 推理子代理独占 GUI，使用实际 macOS Preview 打开上述本次生成 PDF。核心 native/default 各 2 页，以及 headers 各 5 页、image 各 1 页、selected 各 3 页均逐页检查；正常实际尺寸和放大检查“数量”。子代理报告未发现视觉缺陷：正常尺寸“量”上下字腔可辨，表头明显粗于正文，标题无异常轮廓或重影，第二页颜色、粗体和分页正确，补充样例符合预期。

最终逐页记录、截图和窗口关闭情况见 `artifacts/pr108-cross-platform/preview/report.md`。主代理已复核逐页记录、实际尺寸第一页、表头放大及第二页的三张 Preview 截图，确认本轮视觉 PASS。本次打开的 Preview 文档窗口已关闭。

## 范围

功能、自动渲染、PNG 复查和 Preview 均通过。结论仅覆盖固定输入、字体及页面，不覆盖第三方 OFD 阅读器，也不声明等同物理 Bold。既有 Preview Find／原始 Copy 等独立问题不在本轮验收范围。本轮未发布包或创建 Release；验证结论作为 issue #107 的关闭依据。
