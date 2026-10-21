# R14 标准加密有限互操作验收

基线：`ba2d14cc13967441c79a89a19b968c7a9782b756`。上游 Java 固定在 [`84e9eb260ba6fd351ecccf08a01a40305d4af635`](https://github.com/ofdrw/ofdrw/tree/84e9eb260ba6fd351ecccf08a01a40305d4af635)；本次运行使用 Maven `org.ofdrw:ofdrw-crypto:2.4.0`，其运行 JAR SHA-256 为 `4a4098d136c0b597036f8f8904859c31fb4203c5ad5bfa515da2c918d4ce9a54`。检查的 12 个 crypto 源文件与该版本 sources JAR 逐字节相同。上游按 Apache-2.0 提供；此处保留的测试证书及私钥是上游公开的**测试身份**，不得当作可信或生产身份。

## 固定样本

`tests/Ofdrw.Net.Crypto.Password.Tests/Fixtures/standard-java/` 的五份样本均由 Git LFS 管理。`hello.ofd`、`sign_cert.pem`、`sign_key.pem` 来自固定上游源码的测试资源；`java-pass.ofd`、`java-cert.ofd` 由 [`R14Java.java`](../../../e2e/Ofdrw.Net.Crypto.Standard.E2E/R14Java.java) 调用 Java `OFDEncryptor` 后再调用独立的 `OFDIntegrity.protect` 生成。两种输出均有 `alice`、`bob` 两个同类接收人；密码分别为 `first-pass`、`second-pass`。证书样本的两个用户使用该测试证书。它们是测试互操作数据，不构成厂商/标准认证。

| 样本 | SHA-256 |
| --- | --- |
| `hello.ofd` | `1df1080b49774a4084a2b42b18e6cb866eb514d6ad158613c6b389b16eb4136c` |
| `java-pass.ofd` | `86b7d0f2df53a520ba5ab2c5e2bbbd5c00a159fba551d4476f10a968601ccffa` |
| `java-cert.ofd` | `c3953ada0a2dbde3002c69284a2cb113018b0de7a063e333977f327adc7992f7` |
| `sign_cert.pem` | `48c54cf2756a6ac6002a92bc322baab4ada2d051e2a62f7e0cda6ac0a448bbed` |
| `sign_key.pem` | `d3ba4c1a47f8507efd8a8895d6d7f9584771e7c4adbdda065fd6c18a710cec0d` |

运行时使用 `maven:3.9.9-eclipse-temurin-21` 容器、上述 Java JAR 及其传递依赖；`R14Java.java` 的 `enc-pass`/`enc-cert` 生成两份输入，`verify` 调用 Java `OFDIntegrityVerifier`。固定的 LFS 输出让普通 .NET CI 不依赖实时 Maven/Java 下载。重新生成的加密包含随机 FEK、IV，故输出哈希不预期固定。

## 功能与双向交叉检查

- Java → .NET：本包分别以 `bob` 口令和测试证书私钥解出两份 Java 包；按 ZIP 逻辑文件名及每项原始字节与 `hello.ofd` 比对，均相同。两份清单的 Java 签名与精确文件名集合可验，原生清单没有逐文件内容摘要；默认解密拒绝，只有显式 `AllowUnverifiedContents=true` 才返回 `FileContentsVerified=false`。
- .NET → Java：本包从同一 `hello.ofd` 产生口令与证书各含两个接收人的包。Java `OFDIntegrityVerifier` 对两份包均返回 `true`；Java 分别以 `bob` 解密后，原始 OFD 文件名及每项字节与 `hello.ofd` 相同。Java 解密器同时保留旧 `OFDEntries.xml`、`signedvalue.dat`，这两个残留文件已不描述恢复后的明文包，不能作为解密后完整性证据。
- 本包自闭环：签名清单及私有签名逐文件 SM3 扩展均验证。错误口令/证书、密文改动、额外或缺失条目、清单改动拒绝。文件 API 在失败时保持目的文件原值；旧私有 v2 回归继续通过。
- 本地包消费：[`run-standard-crypto-e2e.sh`](../../../scripts/run-standard-crypto-e2e.sh) 生成 Core、Packaging、Crypto.Password 三个实际 `.nupkg`，用隔离 `PackageReference` 消费端验证口令/证书多接收人、内容认证、失败原子性及私有 v2。版本 `0.1.0-r14.local` 仅为本地测试，不是公开发布。包 SHA 与依赖保存在忽略目录 `artifacts/roadmap-v12/R14/package-e2e/package-manifest.json`。

2026-10-21 本次本地复跑：`Ofdrw.Net.Crypto.Password.Tests` **104/104**（含五份已物化 Java fixture、大小写路径更名拒绝）通过；最终解析器修改后的隔离包消费通过，构建 0 warning / 0 error。三个本地包的 SHA-256 依次为 Core `e6ed85dc023947347bf8e7d01b907e9b8742937e859c81fac3b445328cb5f884`、Packaging `7478f549405eb37fd92e12c6fd8d2155986d12b856bba0aa830b8cb8e3d7f727`、Crypto.Password `e33434c4e756e5d69197f80ae3d872308cf31b2546a5ada6784258dad55d32fe`。反向 Java 复跑在本次修改后对两份新 .NET 包的清单均输出 `true`，解密后原始文件条目字节相同；两份 .NET 包的 SHA-256 分别为 `6080c591b5f9ee59cff10014e06679a6440bd2524c99afd5916adbdacce98807`、`13993520eccaa3f0f3d53c5cdbeb53d2a16f5b122811493e6896315d63b8915d`。该哈希仅标识这次随机输出。

本地复查命令：

```bash
export DOTNET_CLI_HOME="$PWD/artifacts/roadmap-v12/R14/dotnet-home"
export DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1 DOTNET_CLI_TELEMETRY_OPTOUT=1
dotnet test tests/Ofdrw.Net.Crypto.Password.Tests/Ofdrw.Net.Crypto.Password.Tests.csproj -c Release --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
./scripts/run-standard-crypto-e2e.sh "$PWD/artifacts/roadmap-v12/R14/package-e2e"
```

## 结论边界

本次代码只改变可选密码包及其文档、测试和 CI，不触及转换、排版或渲染，也没有生成新页面；PNG 视觉复查与 macOS Preview 页面验收对本票不适用。功能/包/双向 Java 检查各自独立记录。签名证书由包外 pin，未验证 PKI 链、吊销或业务身份；Java 原生清单只证明签名文件名，不证明所列文件内容。本包的内容摘要扩展是签名覆盖的私有扩展，Java 验签器不解释它。不得从上述样本推断混合接收人、多层加密、厂商产品互认、GM/T 0099 完整实现、商密认证或长期保存适用性。
