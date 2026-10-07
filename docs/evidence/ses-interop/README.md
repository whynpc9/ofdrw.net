# Issue 12 SES dev/interop 证据

本票新增独立可选包 `Ofdrw.Net.Signatures.SesInterop`。功能范围和失败契约见
[教程](../../tutorials/16-ses-dev-interop.md)。默认依赖及核心
`SupportsBuiltInSesSm2Verification=false` 保持不变；无阅读器互认、法律效力或合格 TSA。

源码候选提交 `b871b19bdbca9f7cc7ba7b9422118e4c0eaf9000`，本地包版本
`0.1.0-ses-interop-b871b19`。

本次功能验证：主会话 35/35 定向、435/435 全套测试；指定 Sol Low 独立实际重跑
35/35 和完整本地包消费 E2E。现有固定 `999.ofd` 仅作 V4 字段检查，SM2 已知答案
来自 BouncyCastle 2.6.2，均没有厂商信任结论。默认 11 产品包仍按原精确 validator
验证，SES 包另行打包；package-only consumer 消费同轮本地包，CLI DLL 与 CLI 包、
对应 SDK 包逐字节匹配。四样例引用完整性为真、默认密码状态 Unsupported/CLI
退出 2、显式固定证书测试验证通过、错误证书失败。

视觉验证：本工作树 `artifacts/ses-interop/round4/files/` 的六份 PDF 各 1–2 页，
共 **12 页**，于 2026-10-08 经 macOS Preview 实际查看。
链路为确定性 `generated-layout.docx`（仅将字体替换为有 OFL 许可的固定 Noto）
→ 显式 Native/默认原生 OFD → SES V1/V4 自签 OFD → PDF → Preview；未用直接
DOCX→PDF 代替。四个自签 PDF 8 页与两个同轮未签基线 4 页的中文、比例英文斜体、
字号/基线、表格蓝底/边框、局部红色粗体、右对齐日期和分页一致；未见缺字、乱码、
裁切、重叠、重影或异常空白页。1×1 SES 测试图片仅存于印章 DER，本票没有可见
签章外观。样例无页眉页脚、页码或正文图片，未对这些内容作新增验收结论。

每份 PDF 均按本轮绝对路径打开，查看两页后关闭，再开下一文件；最后取消自动
打开面板并退出 Preview，用只读进程检查确认退出，已向协调会话释放 GUI。
逐页 PNG 是辅助证据，签名后与各自未签基线像素相等；不能替代上述 Preview。
结论只涵盖这些样例/页面，不表示任意 Word 文档保真，也不提升密码签章效力。

`acceptance.json` 记录基线、源文件冻结清单、环境、许可、模式、实际查看路径、页码、
体积和限制；`manifest.json` 固定所有留存文件。`evidence.tar.zst` 包含同轮 OFD/PDF/
PNG、公共测试证书、SignedValue、包清单/本地候选包、测试日志/TRX 和独立报告。
候选包的 RepositoryCommit 均为上述提交，独立消费前后 SHA-256 未变。
不存在私钥导出。重生签名使用新的随机密钥/nonce/核心声明时间，哈希会变化；布局
内容与断言可重现。此前 IPC、consumer 配置、测试构造歧义和 CLI hash 校验失败日志
保留，不能用失败轮作为验收样例。

提取并逐文件核验（通用安全解包工具的 `--bundle` 参数）：

```sh
python3 scripts/unpack-document-tools-evidence.py artifacts/ses-interop/review \
  --bundle docs/evidence/ses-interop/evidence.tar.zst
```

第三方密码库实际包元数据、许可证和 SHA-256 在 `dependency.json`、`.nuspec`、
`BouncyCastle-LICENSE.md`；固定上游/BC 源码 URL、哈希与核对行号在
`source-manifest.json`。GUI/自动化/独立验证是分别记录的证据，不作互相替代。

PR 的首轮评审、最新 head 复审、CI 及线程关闭状态另行记录；本地通过不表示 PR
已闭环。该票不合并 PR，不发布 NuGet。
