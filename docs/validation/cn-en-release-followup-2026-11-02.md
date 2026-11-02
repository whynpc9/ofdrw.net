# 中英文发布跟进：本地独立核验（2026-11-02）

本轮只验证系统对照来源和字体部署教程，不改变 [既有验收](cn-en-release-local-validation.md)。原 Regular-only “量”字视觉 **FAIL**、原 native/default 页末日期 Preview Find **FAIL** 和原始 Unicode Copy **未完成** 均保留；不宣称发布通过。验证者为 `gpt-6.1-terra`，High。

## 系统对照

对照报告和复现脚本来自 `4a913a3f2626abe92c34a8e602c8565b498c13cd`，详见 [系统原生 PDF 对照](cn-en-preview-control-2026-11-02.md)。独立核对原控制目录 `artifacts/r22-cjk-filled-stroke/` 中七项 SHA-256：

| 文件 | SHA-256 |
| --- | --- |
| `native-font-word-controls/produce.swift` | `ee9443edd9d27210c151247ba50b46f9ced338b7ab8a42b3c3f5bd7541255502` |
| `native-font-word-controls/helvetica-date.pdf` | `2a93dab21e7213bb9b8977921fa0bd38b6b3d028440b70c94790138f0f44c5a3` |
| `native-font-word-controls/helvetica-word.pdf` | `1285418794fd32a9ee7dce68ce589bff4f3219dd93d191c7fa2f1f97995ca774` |
| `native-font-word-controls/preview-review/receipt.json` | `81afd5e5628f6902a81d06f4374ecfd6a632f6f1551074eae8227e7f86ba50ab` |
| `native-font-word-controls/root-closure-audit.json` | `9abf3546cbb0ca10bc9b569735219f57398d445283ff41a4cf034744d1d7cda7` |
| `native-coretext-search-controls-677205e1/produce.swift` | `733db3bcac69596aa74b80b8f974691dd63acf0f5875bdc930862263dfa5d895` |
| `native-coretext-search-controls-677205e1/preview-review/receipt.json` | `015758732ba57e9618cdb9f4e06e351fbccbe23573568f0c316374007ccdaa1e` |

原目录位于 `/Users/wanghongyi/.codex/worktrees/9c7f/ofdrw.net/`。旧真实 Preview 收据中 Helvetica 的 `Alpha` 和 `2026-01-01` 各有一个页面命中，但数值命中数为 null、侧栏 `Page 1 Not found`、Next 禁用，点击结果行无字形对齐高亮。系统 CoreText/CGContext 生成器也出现该现象，说明其并非 SDK 产物独有；机制和 OS 根因仍未确定。本轮未重新打开旧 PDF。对新提取脚本仅执行 `swiftc -typecheck scripts/validation/preview-native-terminal-control.swift`，退出 0；未运行脚本或生成新对照。

## 字体教程

[教程](../tutorials/cn-en-font-deployment.md) 源自 `80ccbfc2917612ee6d07dc3f9bd40500245552d3`，在 `/Users/wanghongyi/.codex/worktrees/c52e/ofdrw-font-guide` 独立验证。教程中的 C# 和项目文件与所运行文件逐字相同；提取的 Python 配方仅有末尾换行差异。CLI 源码静态确认位置参数、`docx-to-ofd`、`ofd-to-pdf`、`--docx-ofd-mode` 和可重复 `--font-directory`。Python fontTools 4.59.2 配方重新运行成功，复用并校验缓存输入，不能视作本轮完整在线下载成功；开发者此前在线取 CJK 许可证时遇到 TLS EOF。

固定变量字体源 SHA-256 `990c807e79c25662a5a9ecf7f971baeb2bf2eab9a559e5ecf15cdfdb8561d21f`；生成清单 SHA-256 `dedd978f69c7e702ac83b6b6ada9a98f4f97474520064492662afa4d11666234`。五个输出身份分别为 CJK Regular `3012a9b63f5eca3e3b38f23a1be5ed504675e394abf8e7a4fa981506582c04aa`、CJK Bold `647a6e375c5a0d8958111173d7b7cec5237e0784cc4490663e43834dfff491be`、Latin Regular `b85c38ecea8a7cfb39c24e395a4007474fa5a4fc864f6ee33309eb4948d232d5`、CJK 许可证 `6a73f9541c2de74158c0e7cf6b0a58ef774f5a780bf191f2d7ec9cc53efe2bf2`、Latin 许可证 `0dab92d0544f7b233403f14b84a663bdbfa746982eda629e7f4f9ffe1b036feb`。

按教程执行 `dotnet build artifacts/font-deployment-app/FontDeployment.csproj -c Release --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false`，使用独立可写 `DOTNET_CLI_HOME`、禁用遥测/首次体验并显式设置既有 `NUGET_PACKAGES`：成功，0 errors、41 warnings，包括已有依赖安全告警。执行教程的 `FontDeployment.dll`，输入 `generated-layout.docx` SHA-256 `17bea68d57776b02e5c68c8915de11b42087809d5310f8fafcbb003db08d7bb6`；显式 native 与默认模式均报告 Native、original text True，无模拟粗体诊断。此结果不是全局字体嵌入保证。

产物在上述教程工作树 `artifacts/font-deployment/output/`：

| 文件 | SHA-256 |
| --- | --- |
| `font-native.ofd` | `e182d7403646a27297c28ce974fd04715b78a51b429c43a4d25a7fcb360f0f92` |
| `font-native.pdf` | `f42a914f38f3f761d4b43de4f5e438a04b43a18c8d55095f6faad3b77d7e1961` |
| `font-default.ofd` | `c9f5c85d865dbbd066a51e59f74bebe286c8a6144cea11c118d969a6976374ac` |
| `font-default.pdf` | `9c8b74a2621426758087ccf96e76fbd91f9ac4d5484ec593075385d56fee9a12` |

两文件各两页，均经 DOCX → OFD → PDF → 144 DPI PNG 和实际 macOS Preview 逐页目视，分别完成 4/4 PNG 与 4/4 Preview；任务文档窗口已关闭。该样例物理 Bold 的“数量”字腔清晰，英文斜体、蓝色表格填充与边框、局部红色粗体及右对齐日期可见，未见明显裁切或样式扩散。此结论只适用于该输入及字体配置。PNG 位于本工作树 ignored `artifacts/cn-en-release/continuation-20261102/png/`；native/default 的第 1 页 SHA 均为 `389b6f4c75e89a02a42905927908dcfeb44ba83c90793a577e41bb0b69deb1ff`，第 2 页均为 `7d9ebcc94bd2d1378cc13116c73abb8c3b87e6c3f0041f4d415ac974d1f666d4`。本轮没有重新做 Copy/Find 验收。

PR91 已由远端合并为 `5d83fb3`；协调者通过连接器比较 `1dcbaaa` 到该合并提交得到空文件差异。本机 Git fetch 因 TLS 失败，尚未在实际合并后 main 上重新构建；上述结果不得冒充该次构建。
