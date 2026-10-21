# SES V5 与 GB/T 35275 有限互操作

`Ofdrw.Net.Signatures.SesInterop` 是**可选包**。它保留既有 V1/V4 测试自签入口，同时提供：

- SES V5 的有界只读字段解析、调用方提供的 V5 印章和 SM2 签名操作、显式完整证书 pin 的双层验签；
- 单签名者、附带内容的 GB/T 35275 `ContentInfo/SignedData` profile，默认对原始 `SM3(Signature.xml)` 签名；独立的 Base64 摘要模式仅作显式兼容选项；
- V5 签名值生成后的时间戳回调。回调收到原始 SM2 签名 DER；返回 token 只会保留在容器内，**不执行 TSA 验证**。

所有入口均须显式构造。核心 `OfdSignatureVerifier` 默认仍只验证 OFD 引用摘要，对这些签名值返回 `Unsupported`；核心 `SupportsBuiltInSesSm2Verification` 仍为 `false`。旧 `SesTestSignatureProvider` 和 `SesTestSignedValueVerifier` 仍只针对本库测试 profile。

## 外部签名操作

调用方实现 `ISesExternalSm2Signer.SignAsync`，以 SM2/SM3、SM2 ID `1234567812345678` 对传入的**准确字节**签名，返回规范 DER `{r,s}`。私钥始终由调用方保存；当前 profile 不定义其他算法或 SM2 ID。调用方需从自己的系统取得证书、V5 印章 DER，并独立核准印章及证书来源。

```csharp
ISesExternalSm2Signer externalSigner = /* HSM or caller-owned signer */;
byte[] signerCertificate = /* DER, independently obtained */;
byte[] makerCertificate = /* seal-maker DER, independently obtained */;
byte[] externalV5Seal = /* prebuilt V5 seal DER */;

var v5 = new SesV5ExternalSignatureProvider(
    signerCertificate, externalV5Seal, externalSigner);
await new OfdSignatureService().SignAsync(unsignedOfd, signedOfd, v5);

signedOfd.Position = 0;
var pinned = new SesV5PinnedVerifier(signerCertificate, makerCertificate);
var report = await new OfdSignatureVerifier(new[] { pinned }).VerifyAsync(signedOfd);
```

V5 的受验子集要求完整证书授权列表、无额外印章扩展或印章时间戳。验签检查文档 `Signature.xml` 的原始字节摘要、包内属性路径、TBS 时间是否落在印章声明的有效区间、签名者证书 pin、印章制作者证书 pin、授权列表、印章签名与文档签名。带 `Z` 的 UTC `SignatureDateTime` 必须与已签 TBS 时间一致；对 Java OFDRW 的无时区格式只验证格式，不推断其 UTC 偏移，也不要求它与独立 TBS 时间相等。签名入口在调用外部签名操作前检查所用时间落入印章有效区间。无法解析的结构、算法、扩展或证书不尝试替代 profile。签名值的外层时间戳在无 TSA 验证器时使显式 V5 验签返回无效；字段读取仍可观察原始 token。

## GB/T 35275 的受验子集

```csharp
var provider = new Gbt35275SignatureProvider(signerCertificate, externalSigner);
await new OfdSignatureService().SignAsync(unsignedOfd, signedOfd, provider);

signedOfd.Position = 0;
var verifier = new Gbt35275PinnedVerifier(signerCertificate);
var report = await new OfdSignatureVerifier(new[] { verifier }).VerifyAsync(signedOfd);
```

此处只实现 `signedData` 内容类型、SM3 摘要、SM2 签名、一个证书和签名者、无属性/CRL、附带 32 字节原始摘要的 profile。`Gbt35275PayloadProfile.Base64Sm3DigestCompatibility` 是单独的兼容入口，签名和验签两侧必须明确选择；不自动尝试两种输入。其他 `SignedData` 组合应报告不支持，不能推断整个 GB/T 35275 标准均已覆盖。

SES V5 与 GB/T 35275 在 OFD `Signature.xml` 中使用相同的 SM2/SM3 OID。核心按 OID 注册一个验签插件；同时使用两种容器时，将各自配置好的 pinned verifier 交给 `SesSm2ContainerVerifier`，它按有界 ASN.1 结构明确分派。若只配置其中一种，另一种返回无效。

## 本次交叉验证的复现入口

`e2e/Ofdrw.Net.Signatures.SesInterop.E2E/CrossInterop/` 生成与验证 .NET 签名值；`JavaInterop.java` 用 `org.ofdrw:ofdrw-sign:2.3.9` 双向生成与验证 V5、GB/T 35275。`GoInterop/verify.py` 用固定版本 `tjfoc/gmsm` 解析两种 GB/T 35275 容器、核对附带的 `SM3(Signature.xml)`，再用 OpenSSL 3 的 SM2 原语独立验签。.NET 输入为同目录 `testdata/r12-Signature.xml`（UTC 时间）；Java 输入为 `testdata/r12-JavaSignature.xml`（上游使用的不带时区格式）。两种输入分别与各自签名值配对验证，避免将 Java 当前生成的 TBS 时间与预先固定的 UTC XML 时间误配。证书、印章和签名值保存在本次忽略的 `artifacts/` 证据目录，不提交测试私钥或重复二进制载荷。

Go 库的 PKCS7 高层 `Verify()` 接受 `.501` 签名算法 OID，而固定 Java GB/T 35275 容器发出 `.301.1`。独立探针显式检查 `.301.1` 后使用 OpenSSL 验证准确的签名内容，不能据此声称 Go 的高层接口原生接受该容器。两种独立路径均只验证这组实测样例。

`FullyValid` 只表示**当前显式注册的容器验签器**与 OFD 引用检查均成功，不能解释为自动证书链信任、吊销状态、印章机构授权、TSA 时间或法律效力。完整证书 pin 是调用方的身份策略输入；本库不从容器证书自动建立信任。标准正文、第三方样本和真实设备的互操作仍须分别核准后才能扩展声称。可见盖章位置由 R11 处理；此处的 `Seal.esl` 不会自行绘制为页面 `StampAnnot`。

参考固定源码：Java [`SESV5Container`](https://github.com/ofdrw/ofdrw/blob/84e9eb260ba6fd351ecccf08a01a40305d4af635/ofdrw-sign/src/main/java/org/ofdrw/sign/signContainer/SESV5Container.java)、[`GBT35275DSContainer`](https://github.com/ofdrw/ofdrw/blob/84e9eb260ba6fd351ecccf08a01a40305d4af635/ofdrw-sign/src/main/java/org/ofdrw/sign/signContainer/GBT35275DSContainer.java)。[GB/T 35275-2017 官方状态](https://openstd.samr.gov.cn/bzgk/std/newGbInfo?hcno=A7B91213CC4862B31BE2C84665CB8F7E)须按实施日期复核；上述源代码不是完整标准正文或认证证据。
