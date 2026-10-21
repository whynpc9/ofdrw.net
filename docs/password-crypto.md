# 可选口令加密包

`Ofdrw.Net.Crypto.Password` 是显式引用的独立 `net8.0` 扩展。原有 `OfdPasswordEnvelope` 继续只处理本包私有 profile v2；新增 `OfdStandardEncryption` 是独立、显式的 Java OFDRW 2.4.0 交换子集。它不进入 `Ofdrw.Net.Converter` 默认依赖，不修改核心 `SupportsGmT0099EncryptionEnvelope=false`。本票未发布 NuGet；本地 feed 消费与交叉样本只证明具名实现/样本。不是商用密码产品或密码认证，没有厂商互认证据，不能当作长期保存格式；OFD-H/档案入库前应解密。

## 显式 Java OFDRW 2.4.0 交换子集（R14）

`OfdStandardEncryption.EncryptAsync` / `EncryptFileAsync` 要求显式传入 1–32 个具名接收人和清单签名回调。接收人可全部为口令，或全部为 SM2 X.509 证书；同一包混合口令/证书、多层加密、逐条目独立 seed 与部分条目选择暂不支持。口令按原始 UTF-8 经 Java SM3 counter KDF 得到 16 字节 KEK，没有盐或可调成本，弱口令可被离线猜测。证书路径用 SM2 C1C3C2 包装 FEK，并按 Java 的 DER `SM2Cipher` 结构存入 `EncryptedWK`。一包只有一个 FEK/IV，SM4-CBC/PKCS7 加密全部原文件条目和映射；这沿用固定 Java 格式，不能把它等同私有 v2 的独立 IV/HMAC 安全属性。

```csharp
var standard = new OfdStandardEncryptionOptions {
    ManifestSigner = new OfdStandardManifestSigner(signerCertDer, SignSm2DigestAsync)
};
standard.Recipients.Add(OfdStandardRecipient.ForPassword("alice", passwordA));
standard.Recipients.Add(OfdStandardRecipient.ForPassword("bob", passwordB));
await OfdStandardEncryption.EncryptFileAsync("source.ofd", "encrypted.ofd", standard);

var restored = await OfdStandardEncryption.DecryptFileAsync(
    "encrypted.ofd", "restored.ofd", OfdStandardCredential.ForPassword("bob", passwordB),
    new OfdStandardDecryptionOptions { ExpectedManifestSignerCertificateDer = signerCertDer });
// restored.FileContentsVerified == true for this package's signed digest extension.
```

`SignSm2DigestAsync` 由调用方实现：输入为 `OFDEntries.xml` 的 32 字节 SM3 摘要，返回对该摘要作 SM3-with-SM2 后的 DER 签名；本包在发布前用证书验证回调结果并写入 GB/T 35275 形状的 `signedvalue.dat`。验签证书必须从包外独立固定；回调/固定证书不验证证书链、吊销、证书用途或业务身份。证书解密还要求包外固定预期接收人证书，并与 `UserCert` 字节匹配。

`OFDEntries.xml` 的标准 `FileList` 只签署文件名清单。本包在同一个签名下增加 `urn:ofdrw-net:standard-encryption-content-digests:1` 扩展，逐一记录所有**存储中的文件条目**（包括加密元数据和密文）的 SM3；解密前验证签名、路径集合闭合和每项字节摘要。Java 的验签器可验这份清单并拒绝额外条目，但不检查该扩展的内容摘要。固定 Java 自身生成的清单没有逐文件摘要：默认解密拒绝；只有调用方显式设置 `AllowUnverifiedContents=true` 才恢复，结果 `FileContentsVerified=false`。这不是内容认证成功，CBC 填充检查也不是完整性证明。清单签名、路径闭合、逐文件内容和 PKI 信任必须分别报告。

固定 Java 2.4.0 写端实际省略 seed 的 `EncryptCaseId`；本包读取时在调用方已显式选择口令或证书且字段形状吻合时接受这个缺省，写端在同类接收人包中明确写 `1.1.1` 或 `1.1.2`。Java 对 `Report.dat` 等条目可能把密文写到仅大小写不同的 `report.dat`；本包允许**同一映射**的这种别名，恢复原始路径大小写，但拒绝与其他条目的碰撞。解密拒绝未加密包、未知/重复接收人、缺失密文、额外文件、重名/非规范路径、签名或摘要不符；不会复制 Java 的“缺失密文跳过”和“映射解密失败改试明文”回退。与 Java 互操作的实际样本及包身份见 [R14 验证记录](evidence/r14-standard-encryption/README.md)。Java 解密 .NET 包能恢复原始文件条目，但其自身保留已经指向密文条目的旧 `OFDEntries.xml`/`signedvalue.dat`；调用方不得把该残留清单作为恢复后包的有效完整性声明。

私有 v2 与标准入口不能自动互解或降级；下文的 v2 说明仅适用于 `OfdPasswordEnvelope`。

```csharp
using Ofdrw.Net.Crypto.Password;
var options = new OfdPasswordOptions();
options.PageIndices.Add(1); // 第二页，零基
// 可显式补选共享资源；必须是存在的规范 ZIP 条目名。
options.EntryNames.Add("Doc_0/Res/image.png");
await OfdPasswordEnvelope.EncryptFileAsync("source.ofd", "encrypted.ofd", password, options, token);
await OfdPasswordEnvelope.DecryptFileAsync("encrypted.ofd", "restored.ofd", password, cancellationToken: token);
```

使用完整内存结果时，调用 `EncryptAsync(Stream, password, options, token)` 或 `DecryptAsync(Stream, password, options, token)`；它们保持输入流打开，返回完整 `byte[]`，失败不返回部分明文。调用方负责清理返回的明文。没有默认 CLI 命令。

## 筛选与保留

必须显式选择至少一个条目或页。`EntryNames` 是精确的规范包内路径，禁止 `/` 开头、反斜杠、空段、`.`/`..`、冒号和控制字符。`PageIndices` 按唯一 DocBody 的 Pages 顺序，以零为第一页，只选择 `Page@BaseLoc` 指定的内容文件。未选页与选中页共享 BaseLoc 时拒绝。

页选择不会自动加密字体、图片、模板、附件、注释或其他地方的同文内容，因此不是整页信息的保密闭包。调用方需显式补选资源。可把全部原始 ZIP 文件条目加入 EntryNames，包括 OFD.xml，来选择全包内容；加密元数据仍明文可见。

普通未知条目按字节保留，解密恢复原始文件条目的名称和载荷。ZIP 的压缩方式、时间戳、文件权限和目录占位项不保留；输出不压缩，因此通常明显变大。只支持单 DocBody、单层、单用户。已有加密元数据、保留目录 `PasswordCrypto`、签名声明或约定 Signs 载荷一律拒绝；需要处理签名时先由调用方显式清签。不自动清签或保证现有签名继续有效。

## 算法与结构

v2 的口令派生采用 [RFC 7914 scrypt](https://www.rfc-editor.org/rfc/rfc7914)，后端声明精确 NuGet 依赖范围 `[2.6.2]`：`BouncyCastle.Cryptography 2.6.2`（MIT）。消费应用显式覆盖后端版本不属于已验证的契约；NuGet 的直接依赖优先规则及其版本约束警告不能视为本包验证了其他版本。口令严格 UTF-8，不 trim、不规范化；接受 1–1024 UTF-16 单元，拒绝无效代理项。

每包用系统随机源生成独立 32 字节盐，以固定 `N=131072, r=8, p=1, dkLen=48` 派生：前16字节为包装密钥，后32字节为 HMAC-SHA256 密钥。主 ROMix 表约128 MiB；本包按160 MiB保守预算，工作单位为 `N*r*p=1048576`。参数是本私有版本的策略，并非 RFC 对所有环境规定的统一安全等级；口令仍可能被离线猜测，盐和成本使每次猜测必须独立执行昂贵派生。此工作因子不保证弱口令无法恢复。

每次独立随机生成16字节FEK、包装IV、每个载荷IV与映射IV。使用SM4-CBC/PKCS#7包装FEK并加密选定载荷和映射；每个载荷/映射记录为16字节IV前缀加完整CBC密文。不会跨条目复用同一FEK/IV组合。后端算法实现不在本库重写。

`Encryptions.xml` 采用 `urn:ofdrw-net:password-profile:2` 的 `Version="2"`，包含唯一该命名空间 `Kdf` 元素及 `Algorithm="scrypt", N="131072", r="8", p="1", dkLen="48", Salt="base64"`。参数必须精确匹配；符号、空白、前导零、重复/外来元素、未知算法、短盐或恶意超大成本均在后端分配前拒绝。`EncryptInfo ID="1"` 的Provider为 `Ofdrw.Net.Crypto.Password`、Company为 `Ofdrw.Net`、Version为2，范围Partial表示显式条目选择，固定引用 `/PasswordCrypto/decryptseed.dat` 与 `/PasswordCrypto/entriesmap.dat`。seed根仍拼写为 `DecyptSeed`，但 `EncryptCaseId="2.scrypt-sm4-etm"` 明确标识私有套件；含唯一UserInfo/EncryptedWK/IVValue和空ExtendParams。映射根 `EncryptEntries ID="1"` 同样要求v2版本，映射使用Path/EPath，Inventory继续包含全部原条目的Path、Length和SM3以验证精确恢复。

固定条目 `PasswordCrypto/auth.dat` 保存32字节HMAC-SHA256。认证输入为UTF-8域 `Ofdrw.Net.Password.Profile.v2` 加一个零字节、8字节大端逻辑文件数，再按条目名Ordinal顺序依次加入8字节大端路径字节长度、严格UTF-8的实际路径名、8字节大端内容长度和原始内容字节。只排除精确拼写的auth.dat自身，其余所有文件条目（包括未加密条目、入口/seed原始XML字节、IV、包装、映射与载荷密文）均认证。路径大小写变更、XML格式/注释改写、缺失/多余条目都会拒绝；auth条目须精确固定拼写。使用固定时间比较，在解开FEK、CBC解密、映射明文解析和输出前先验MAC。错误口令与认证失败不区分。

这是私有encrypt-then-MAC格式，设计顺序参考[既有EtM原则](https://www.rfc-editor.org/rfc/rfc7366#section-3)，不声称是RFC格式或GM/T完整性协议。持有口令者可以重新构造有效MAC，因此不提供签名身份、作者来源或防回放；不能检测攻击者把整个包换成另一个合法包。认证对象是逻辑文件名/载荷，不是ZIP物理表示：压缩方式、时间戳、注释、零载荷目录占位不纳入MAC，也不保留。库存SM3继续作为已认证恢复结构的额外字节检查，不代替MAC。

新写入和读取仅支持v2；旧v1、缺失/未知/混合版本全部拒绝，无兼容开关和弱化回退。历史v1字节不会被重新解释或追溯增强；若需恢复旧文件，应在受控的旧版本工具中先解密，再显式用v2重加密。历史验收目录保留v1身份，不能作为v2安全或视觉证据。固定上游[ofdrw 5fe9c427](https://github.com/ofdrw/ofdrw/tree/5fe9c4276c64e40b455e6ea649b695adf8a9a734/ofdrw-crypto)仅提供早期结构轮廓：v2与其KDF、共享IV及无MAC设计明确不兼容。证书、厂商互认、通用GM/T支持、多重加密及多人访问控制均不在范围。

## 预算、失败与清理

默认压缩输入 64 MiB、单条目展开 16 MiB、总展开 64 MiB、10,000 个条目/页、元数据 4 MiB、结果 ZIP 配置上限80 MiB（实际取输出上限与输入上限的较小值，默认有效64 MiB，含ZIP头/中央目录/结束记录）。压缩比预算必须是至少1的有限值，以便不压缩的输出能被同组选项重新加载。原始输入从调用方当前Position一次有界快照到EOF，压缩输入上限须可由内存数组表示（不大于int.MaxValue）。解压前检查原始ZIP名称，禁止反斜杠/空白文件basename/大小写重复；目录仅允许单尾斜杠且零载荷，避免共享Loader规范化或跳过载荷造成静默改名/丢失。快照复制缓冲、扩容旧buffer和最终buffer均清理；调用方流保持打开且不清空其原始buffer。沿用 OfdPackageLoadOptions 的实际展开与压缩比限制，并检查密文膨胀、映射/XML和输出预算。输入/输出同路径拒绝；父目录必须存在。可调预算需考虑本包在内存中同时持有输入、密文、恢复数据及 ZIP 的峰值。

错误口令、MAC失败、损坏的包装/映射/载荷或恢复校验失败抛 InvalidDataException；不声称能区分错误口令与损坏。非本包 profile、未知结构、缺失/多余/冲突条目、签名和预算超限也明确失败。取消为 OperationCanceledException，I/O 错误保持原始 I/O 异常。

KDF预算独立于ZIP预算：`MaxKdfMemoryBytes`默认160 MiB，`MaxKdfWorkUnits`默认1048576，低于该版本要求时拒绝，增加预算不会提高或降低固定参数。单个已加载程序集实例以异步gate串行KDF及其变换，排队可取消；它不限制排队调用已经加载的ZIP总内存，也不覆盖多个程序集加载上下文。BC scrypt没有中途取消接口：后端调用期间取消会等待本次有界计算结束，再清理并抛取消，绝不后台脱离执行或返回部分结果；不保证硬实时取消。SCrypt主工作数组按后端finally清理，不扩大为所有内部副本的不可恢复擦除保证。

文件 API 先完成恢复和校验，再 CreateNew 创建目标同目录随机 `.ofd-password-*.tmp`；Unix 权限 0600，Windows 继承目录 ACL。关闭文件并最后检查取消，随后一次 rename 覆盖目的文件；提交前失败保持旧目的文件或保持目的文件不存在。提交后为成功边界，不再因后到取消返回伪失败。

成功返回的输出保证符合相同预算的输入大小；预算不足时加密失败，不返回无法用同组选项回读的包。生成XML byte[]先登记所有权再判断metadata预算，超限也清理数组；序列化峰值分配不据此宣称受metadata硬上限约束。finally 清理本库拥有的密钥、口令 UTF-8 与已取得的工作缓冲，删除暂存文件；清理失败为可观察的 I/O 异常。加密不向临时目录展开原包，解密只在全部校验成功后暂存完整结果。共享loader尚未返回的中途失败缓冲、托管字符串/后端内部缓冲、不可恢复擦除、进程被终止或文件系统故障不属于清理保证。恢复输出本身是调用方要求的明文，需按调用方的数据策略管理。

## 验证边界

固定向量、损坏/错误口令/取消原子失败、部分页/全条目恢复、可选依赖及实际页面证据见 [11 票记录](evidence/password-crypto/README.md)。只有该记录中实际测试的样本/页面可以作验收结论，不推断任意复杂文件保真或外部阅读器互认。
