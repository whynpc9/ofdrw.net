# R14 标准加密交换的有限互操作设计

状态：2026-10-21 实施前的只读设计检查点；下文的“待验证”仅记录当时状态，运行结论见 [R14 验证记录](evidence/r14-standard-encryption/README.md)。基线 `ba2d14cc13967441c79a89a19b968c7a9782b756`，上游 Java 固定 `84e9eb260ba6fd351ecccf08a01a40305d4af635`。R00 契约已合并。本设计先固定可证实的格式与失败语义，再实施。

## Java 字节契约核查

| 部分 | 上游固定源码 | 可证实的编码/行为 |
| --- | --- | --- |
| 容器 | `ofdrw-crypto/OFDEncryptor.java`, `OFDDecryptor.java` | `Encryptions.xml` 记录 `EncryptInfo`、Provider、`EncryptScope=All`、seed/map 的绝对路径。单个 FEK 16 字节及单个 IV 16 字节作用于全部所选条目和加密 map；SM4-CBC/PKCS7，密文不含 IV 前缀。映射为加密的 `EncryptEntries` XML，每项 `Path`/`EPath`。输出已删除对应明文。解密器遇到缺失密文会跳过，map 解密失败会尝试把密文当明文解析；.NET 不复制这两种回退。 |
| 口令接收人 | `enryptor/UserPasswordEncryptor.java`, `decryptor/UserPasswordDecryptor.java`, `ofdrw-gm/support/KDF.java` | 原始口令 UTF-8，经 SM3 counter KDF 派生 16 字节 KEK；用公共 IV 做 SM4-CBC/PKCS7 包装 16 字节 FEK，`EncryptedWK` 为 32 字节。`UserInfo` 有 `UserName`、可选 `UserType`、`EncryptedWK`、`IVValue`。无盐和口令成本参数；口令路径无法等同私有 v2 的 scrypt 防离线猜测能力。 |
| 证书接收人 | `enryptor/UserCertEncryptor.java`, `decryptor/UserCertDecryptor.java`, `ofdrw-gm/sm2strut/SM2Cipher.java` | 从 X.509 证书取 SM2 公钥，以 SM2Engine C1C3C2 包装 16 字节 FEK；`EncryptedWK` 为 DER `SEQUENCE{X INTEGER,Y INTEGER,hash OCTET STRING,cipherText OCTET STRING}`，并存 DER `UserCert` 与公共 IV。解包本身不建立证书链或撤销信任。 |
| 多接收人 | `OFDEncryptor.java` | 对接收人列表循环写多个 `UserInfo`，共用 FEK/IV；Java 写端计算 `encryptCaseId` 但没有调用 `setEncryptCaseId`，实际 seed 可能缺该属性，Java 读端显式容忍缺失。`DecyptSeed.setEncryptCaseId(String)` 自身使用更新属性方法而非新增。实现须先用真实样本确认输出，不把此缺陷推成标准要求。 |
| 防夹带 | `integrity/OFDIntegrity.java`, `OFDIntegrityVerifier.java`, `core/integrity/OFDEntries.java` | 独立生成 `OFDEntries.xml`：`DocEntries`/`FileList` 内各 `File` 只载 ID 和 FileLoc，`SignedValueLoc` 指向 `signedvalue.dat`；签名值是清单字节的 GB/T 35275 ContentInfo。Java 验证器检查签名及包内额外条目，但没有逐文件摘要，也没有查清单中列出的文件都存在。.NET 将额外与缺失都拒绝。 |

上述相对 Java 路径位于固定 SHA 的 `ofdrw-crypto/src/main/java/org/ofdrw/crypto/` 等相应目录。公开实现是一个交叉实现目标，不单凭类名或注释证明任何 GM/T 0099 认证、厂商交换或 PKI 信任。

## API 与身份边界

- 保留 `OfdPasswordEnvelope` 私有 v2、原默认 API 和原格式。新增独立 `OfdStandardEncryption` 显式入口及选项；不能由 v2 自动探测/回退至无 MAC 格式。
- 加密选项明确列出接收人（口令或 SM2 X.509 DER）；一次操作至少一个，用户名同一大小写规则下不得重复；先支持同一种方案的多个接收人。混合口令/证书、多层、多个 EncryptInfo、每条目独立 seed、非标准/歧义 XML 拒绝，直到双方样本可证明安全匹配。证书私钥输入仅用于解密，不表示链信任。
- 为 Java 原样输出提供单独声明的兼容读取：允许 seed 缺 `EncryptCaseId`，但接收人种类必须由调用方显式指定并与证书字段/包装长度相符；不猜测或尝试所有用户。写端在单一方案下填写 `1.1.1` 或 `1.1.2`，须验证 Java 能读。
- ZIP 路径、大小写重复、预算、单 `DocBody`、签名包显式清理和文件原子替换延续 v2 的保守界限。拒绝输入已加密、部分密文缺失、明密路径冲突、未知压缩/路径语义和超预算。需与 Java 实际产生的条目命名兼容，不强制其 Provider 名等于 .NET。

## 完整性与允许/拒绝矩阵

| 输入/操作 | 预定结果 | 可声称的验证 |
| --- | --- | --- |
| v2 原包与原 API | 原行为不变 | v2 HMAC 对全部逻辑文件名/字节认证。 |
| 标准兼容包，唯一指定用户且 map/条目完整 | 解密并在全部检查后一次发布 | 仅能证明解密与结构；SM4-CBC/PKCS7 不提供主动篡改检测。 |
| 有有效 `OFDEntries.xml` 与 pinned signer | 清单签名验证且路径集合精确一致 | 证明清单签名及无额外/缺失条目；**不证明文件内容**。独立报告 signer pin 和 PKI 信任状态。 |
| 内容字节被改变而清单不变 | 不把清单成功当作内容成功；若受保护内容验证不存在，返回明确的 `unverified` 状态/拒绝要求内容认证的 API | Java 清单机制本身不含逐文件内容摘要。若业务要求防主动篡改，必须另用可信逐文件摘要/签名或私有 v2；不能无声成功。 |
| 错误身份、重复/歧义接收人、损坏包装、非法路径、缺失/额外条目、清单签名失败 | fail closed，无输出/旧文件不变 | 解密与清单检查各自的失败。 |

需确定 API 是否让无内容认证的标准格式以明确结果类型恢复，还是只允许调用者显式选择其风险；不能把 `DecryptAsync` 的普通成功误解为认证。设计审查与真实 Java 样本之后固定这一点。

## 实施与验收门

1. 在实际 Java 环境用固定源码生成口令、证书、多同类接收人输入；记录源码/fixture SHA、许可证与生成命令。反向用 .NET 输出由 Java 解密，逐条比对原包逻辑文件名/字节。当前机器无本机 JDK/Maven，已确认 Docker 可用；Java 执行门尚未完成。
2. 验证 wrong credential、map/密文、额外/缺失/重复路径、清单签名、受保护内容篡改的实际结果；内容完整性结论与清单结论分列。不能证明的要求保持阻断，不用静态推理替代交叉测试。
3. 定向测试、原 v2 回归、本地真实包消费、PR 当前 SHA CI/review。若使用新生成页面，经本次实际 OFD/PDF 的逐页 PNG 与 Preview 验收；不复用历史视觉结论。
