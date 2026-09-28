# GM/T 0099 口令加解密（可选扩展包）

**What to build:** 独立可选包：用口令包装文件密钥，SM4 加密选定条目，自加密自解密闭环，可只加密部分页。核心 SDK 能力标志不变；不得当成长期保存格式或「核心已支持加密」。

**Blocked by:** None (can start immediately)

**Priority:** P2

**Status:** ready-for-agent

- [ ] 自加密自解密闭环；可只加密部分页
- [ ] 不进入 Converter 元包；不改核心 `SupportsGmT0099EncryptionEnvelope`
- [ ] 文档写明无厂商、非商用密码产品、非长期保存件

## Parent

[docs/capability-roadmap.md](../../../docs/capability-roadmap.md) P2-01

## What to build

没有密码厂商时，能做的上限是结构正确的口令密文包。这张票做一个默认关闭的扩展：写出 `Encryptions.xml`、口令派生的文件密钥、条目密文，并能用同一口令解开。调用方可只加密部分页。

禁止：打进默认 Converter 元包；把核心 `OfdCryptographicCapabilities.SupportsGmT0099EncryptionEnvelope` 改成 true；宣称商用密码产品或档案长期保存（OFD-H / 档案方向要求入库前解密）。

证书加密、完整性协议、多重加密不在这张票（需明确调用方后再做）。

## Acceptance criteria

- [ ] 可选包能加密再解密同一份样例，明文页内容恢复；可配置只加密部分页/条目
- [ ] 核心包能力标志保持 `SupportsGmT0099EncryptionEnvelope = false`；扩展用自己的类型声明能力
- [ ] 不进入 `Ofdrw.Net.Converter` 默认依赖
- [ ] 文档写明：无厂商、不是商用密码认证、不能当长期保存格式
- [ ] 错误口令失败可观察，不写出看起来像明文的损坏包冒充成功

## Blocked by

- None (can start immediately)
