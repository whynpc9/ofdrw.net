# Issue 12 SES dev/interop 证据

本票新增独立可选包 `Ofdrw.Net.Signatures.SesInterop`。功能范围和失败契约见
[教程](../../tutorials/16-ses-dev-interop.md)。默认依赖及核心
`SupportsBuiltInSesSm2Verification=false` 保持不变；无阅读器互认、法律效力或合格 TSA。

源码候选提交 `92073352b33b2b26804e6d033cefa4a48fe31563`，本地包版本
`0.1.0-ses-interop-9207335`。

本次功能验证：主会话 37/37 定向、437/437 全套测试；指定 Sol Low 独立实际重跑
37/37 和完整本地包消费 E2E。现有固定 `999.ofd` 仅作 V4 字段检查，SM2 已知答案
来自 BouncyCastle 2.6.2，均没有厂商信任结论。默认 11 产品包仍按原精确 validator
验证，SES 包另行打包；package-only consumer 消费同轮本地包，CLI DLL 与 CLI 包、
对应 SDK 包逐字节匹配。四样例引用完整性为真、默认密码状态 Unsupported/CLI
退出 2、显式固定证书测试验证通过、错误证书失败。

本轮视觉验证 **12/12 页通过**：2026-10-08 在 macOS Preview 实际重新打开
`artifacts/ses-interop/round5/files/` 的 native-v1/native-v4/default-v1/default-v4
和 baseline-native/baseline-default 六份 PDF，各查看第 1–2 页，并逐份核对
绝对 Window URL、页号与已冻结 SHA-256。链路为许可字体确定性 DOCX → 显式
Native/默认原生 OFD → SES V1/V4 自签 OFD → OFD-to-PDF → Preview。
中文和比例英文、斜体标题、蓝色标题、表格蓝底/边框、第二页局部红色粗体、右对齐
日期和分页与同轮未签基线一致，未见缺字、乱码、裁切、重叠、重影或异常空白页。
PNG 相等仅作辅助；本次结论来自本轮 PDF 实看，未移用旧轮验收。

每份 Round5 PDF 在查看第二页后立即关闭，最后核对 Window 菜单无本票自签文件；
同名残留 baseline-default.pdf 的 URL 属于 587a/password-crypto/review3-current，
并非本票。无本票 Open/GoTo 面板残留，GUI 已释放。开场出现的已释放旧 PNG 集合
及旧 baseline-native.pdf 曾清理，来源未确认的 PostScript 临时文档及其他非任务
窗口未操作，Preview 留在运行状态；不宣称关闭了无关用户文档。
样例无正文图片、页眉页脚或页码，本票不生成可见盖章外观，也不提升密码效力。
此前锁屏 pending 记录保留在包内 `history/round5-preview-pending.json`。

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
