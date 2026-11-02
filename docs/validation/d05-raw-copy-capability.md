# D05：原始 Unicode Copy 能力确认

日期：2026-11-02（Asia/Shanghai）。结论：**inconclusive／未完成**。本轮未取得实际 Copy 的原始 Unicode，不能判定复制保真通过或存在产品缺陷。

## 基线与能力证据

从干净工作树、已 fetch 的最新 `origin/main` 开始，源码基线为 `259b7fbf383f69b4b7a9fc0df59d87b6c680603e`（含 D04／PR100），分支 `codex/d05-raw-copy-evidence`。本轮新增验收记录并更新路线图链接，不修改产品、字体、样例、发布标签或历史验收结论。

开发协调为 gpt-6.1-sol High；gpt-6.1-terra High 独立执行本地能力确认，独占 GUI 资源。Terra 首次调用 `cua.getState()`，读取当前工具返回的 API 文档与应用清单，随后于 2026-11-02 22:13 CST 记录能力检查结果。未绑定或操作任何应用，随后释放 GUI。

当前文档提供 App 的 `paste(text, options)`、`pressKey`、`typeText`、AX／截图／选择等操作，**没有提供原始 clipboard／pasteboard 读取接口**。`paste(text, options)` 在 macOS 使用系统剪贴板输入调用方给定的文本并恢复先前内容，不能取得实际 Copy 的结果。工具同时要求：

> Use `cua_repl` (JavaScript) for all UI actions

> Do not use other technologies besides `cua_repl` for computer interactions, unless specifically requested by the user

本任务不授权绕过 CUA／平台限制；没有调用 `pbpaste`、AppleScript、NSPasteboard 或其他原生剪贴板旁路，也没有用 PDF／AX 文本提取或粘贴到搜索字段替代原始 Copy。已有剪贴板未读取、未保存；本次未执行 Copy，原始数据状态为 **NOT_READ**。该结论只描述本轮接口能力，不推断未来工具或其他人工环境同样无法读取。

## 本轮验收范围

能力确认在 45 分钟主动验证上限内结束。因前置读取能力缺失，没有为本轮重新构建或生成 OFD／PDF，也未复用旧产物充当本次验收。

| 项目 | 结果 |
| --- | --- |
| 原始 Unicode Copy | 未完成；无受支持且可安全归属本次实际 Copy 的读取路径 |
| 实际复制范围／页码／原始 Unicode | 无；未执行 Copy，未取得数据 |
| 本次 OFD／PDF、产物 SHA-256 | 无；能力前置不满足，未生成产物 |
| 产品功能测试／本地包消费 | 未运行；仅记录工具能力，不改产品代码 |
| 自动渲染／逐页 PNG 视觉复查 | 未执行，0 页；本轮无生成页面 |
| macOS Preview 页面验收 | 未执行，0 页；未打开任务窗口 |
| Preview Find／其他阅读器 | 未重测，不作通过结论 |

## 必要下一步与判定规则

只有获得受支持的原始剪贴板读取接口，并能安全确认内容来自本次任务的实际 Copy 后，才重开限定补验。读取前不得触及用户已有的无关剪贴板内容。继续使用[真实 Regular／Bold profile](../release-profile-cn-en.md)及未修改 F01，由当次准确源码生成显式 native 和默认模式 OFD→PDF→macOS Preview；记录源码、输入、字体与产物 SHA-256、实际页码和选中范围。

预先声明的换行处理为：先完整保存原始 Unicode 序列与码点；逐字比较时仅允许将 CRLF／CR 统一为 LF，并分别报告原始相等和换行统一后相等。不得删除换行、折叠空格、裁剪首尾空白或进行 NFC／NFKC 等 Unicode 归一化来制造相等；若所选范围包含阅读器自动插入的换行，另列差异并解释，不将其静默忽略。此规则尚未执行，不是本轮复制成功的证据。

[历史本地验收](cn-en-release-local-validation.md)、[preview.10 发布记录](preview10-publication-2026-11-02.md)及已发布标签保持原样。原 Regular-only 视觉 FAIL、Preview Find 异常、原始 Unicode Copy 未验证继续保留；本轮不新增已证实的复制缺陷，也不称整体视觉或发布门通过。
