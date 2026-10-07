# 可选口令加密包

`Ofdrw.Net.Crypto.Password` 是显式引用的独立 `net8.0` 扩展，提供本包私有 profile v1 的自加密、自解密。它不进入 `Ofdrw.Net.Converter` 默认依赖，不修改核心 `SupportsGmT0099EncryptionEnvelope=false`。不是商用密码产品或密码认证，没有厂商互认证据，不能当作长期保存格式；OFD-H/档案入库前应解密。

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

算法依据固定上游 [ofdrw 5fe9c427](https://github.com/ofdrw/ofdrw/tree/5fe9c4276c64e40b455e6ea649b695adf8a9a734/ofdrw-crypto)。后端 `BouncyCastle.Cryptography 2.6.2`，MIT。

口令严格 UTF-8，不 trim、不规范化；接受 1–1024 UTF-16 单元，拒绝无效代理项。按 GB/T 32918.3 §5.4.3 的 SM3 计数器 KDF（大端计数器从 1 开始）取 16 字节 KEK。每次用系统随机源产生独立 16 字节 FEK 与 IV；SM4-CBC/PKCS#7 包装 FEK，包装密文 32 字节。选定载荷及映射表使用 FEK 加密，删除相应明文条目。

这不是慢口令 KDF：无盐、无工作因子。同一操作的载荷/映射表/包装沿用上游的共同 IV 轮廓；同 FEK/IV 多条目会泄露相同明文前缀。没有 MAC 或认证加密，也不提供抗主动攻击者整体替换/重构包的保证。

入口 `Encryptions.xml` 包含唯一 `EncryptInfo ID="1"`，Provider 名为 `Ofdrw.Net.Crypto.Password`、版本 1，范围 `Partial` 表示显式条目选择。它引用 `/PasswordCrypto/decryptseed.dat` 与 `/PasswordCrypto/entriesmap.dat`。前者是明文 XML，沿用固定上游实际拼写 `DecyptSeed ID="1" EncryptCaseId="1.1.1"`，含唯一 UserInfo/EncryptedWK/IVValue 和空 ExtendParams。后者是 SM4 密文，解密根为 `EncryptEntries ID="1"`，映射用 `Path`/`EPath` 绝对包内路径。

入口与映射均含 `urn:ofdrw-net:password-profile:1` 的 Version=1。映射还包含该私有命名空间下的 Inventory：全部原始条目的 Path、Length 和 SM3。解密核对输入条目精确集合、全部恢复长度/摘要后才返回或写出。这是满足损坏失败要求的恢复字节故障检测，不是 GM/T 完整性协议、MAC 或认证。它与上游结构轮廓有私有扩展差异；没有该版本/清单的上游或厂商包明确拒绝，不做弱化回退。恢复校验覆盖原始文件条目及规定的结构约束，不保证检测 ZIP 元数据或未纳入库存的描述字段的逐位变化。证书、完整性协议、多重加密和多人访问控制不在本包范围。

## 预算、失败与清理

默认压缩输入 64 MiB、单条目展开 16 MiB、总展开 64 MiB、10,000 个条目/页、元数据 4 MiB、结果 ZIP 配置上限80 MiB（实际取输出上限与输入上限的较小值，默认有效64 MiB，含ZIP头/中央目录/结束记录）。压缩比预算必须是至少1的有限值，以便不压缩的输出能被同组选项重新加载。沿用 OfdPackageLoadOptions 的实际展开与压缩比限制，并检查密文膨胀、映射/XML和输出预算。输入/输出同路径拒绝；父目录必须存在。可调预算需考虑本包在内存中同时持有输入、密文、恢复数据及 ZIP 的峰值。

错误口令、损坏的包装/映射/载荷或恢复校验失败抛 InvalidDataException；不声称能区分错误口令与损坏。非本包 profile、未知结构、缺失/多余/冲突条目、签名和预算超限也明确失败。取消为 OperationCanceledException，I/O 错误保持原始 I/O 异常。

文件 API 先完成恢复和校验，再 CreateNew 创建目标同目录随机 `.ofd-password-*.tmp`；Unix 权限 0600，Windows 继承目录 ACL。关闭文件并最后检查取消，随后一次 rename 覆盖目的文件；提交前失败保持旧目的文件或保持目的文件不存在。提交后为成功边界，不再因后到取消返回伪失败。

成功返回的输出保证符合相同预算的输入大小；预算不足时加密失败，不返回无法用同组选项回读的包。finally 清理本库拥有的密钥、口令 UTF-8 与已取得的工作缓冲，删除暂存文件；清理失败为可观察的 I/O 异常。加密不向临时目录展开原包，解密只在全部校验成功后暂存完整结果。共享loader尚未返回的中途失败缓冲、托管字符串/后端内部缓冲、不可恢复擦除、进程被终止或文件系统故障不属于清理保证。恢复输出本身是调用方要求的明文，需按调用方的数据策略管理。

## 验证边界

固定向量、损坏/错误口令/取消原子失败、部分页/全条目恢复、可选依赖及实际页面证据见 [11 票记录](evidence/password-crypto/README.md)。只有该记录中实际测试的样本/页面可以作验收结论，不推断任意复杂文件保真或外部阅读器互认。
