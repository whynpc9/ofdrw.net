# SES V1/V4 dev/interop 测试容器

`Ofdrw.Net.Signatures.SesInterop` 是独立可选包，支持常见 SES V1/V4 字段检查和
SM2 测试密钥的自签自验。**仅供 dev/interop**：没有阅读器互认、法律效力、
证书链/吊销校验、商用密码认证、印章外观校验或合格 TSA。生成的自签证书和
印章均标注 `OFD SES DEV INTEROP TEST`，不仿制厂商或税控印章。

本包不由 Core、Writer、Converter 或默认 CLI 依赖，也不会自动注册验证器。
核心 `SupportsBuiltInSesSm2Verification` 保持 `false`；本包自身的
`SesInteropCapabilities` 只描述结构检查和显式测试自签能力。

## 显式使用

```csharp
using Ofdrw.Net.Signatures.Signing;
using Ofdrw.Net.Signatures.Verification;
using Ofdrw.Net.Signatures.SesInterop;

var identity = SesTestIdentity.Generate(); // 每次生成新的临时测试密钥；不提供私钥导出
var provider = new SesTestSignatureProvider(identity, SesVersion.V4); // 或 V1
await new OfdSignatureService().SignAsync(unsignedOfd, signedOfd, provider);

signedOfd.Position = 0;
var ordinary = await new OfdSignatureVerifier().VerifyAsync(signedOfd);
// 引用摘要可以有效；密码状态 Unsupported，FullyValid=false。

signedOfd.Position = 0;
var verifier = new SesTestSignedValueVerifier(identity.CertificateDer);
var testReport = await new OfdSignatureVerifier(new[] { verifier }).VerifyAsync(signedOfd);
// 只有显式注册且两层签名、测试配置和引用均匹配时，testReport.FullyValid=true。
// 这只是调用方选择的测试配置通过，不表示法律/证书信任/厂商互认。

var fields = SesSignedValueReader.Parse(signedValueDat);
Console.WriteLine($"SES {fields.Version}: {fields.Seal.Name}, {fields.PropertyInformation}");
// Parse 只检查结构；不能以字段存在推断签章可信。
```

`CertificateDer` 必须由调用方在包外固定。V4 外层证书在 TBS 之外；验证器比较
完整 DER，防止同公钥但不同证书的替换。印章授权列表和制章者也须使用同一证书。
签名服务不会向页面添加可见盖章图像；1×1 测试图片只在 SES 印章数据中。

默认 CLI 没有加载此包或注册测试验证器的开关：

```sh
ofdrw verify-signatures self-signed-v4.ofd
# reference integrity valid; signed-value algorithm requires a registered verifier.
# exit code 2；不输出 fully valid。
```

## 容器与密码约定

| 项目 | V1 | V4 |
| --- | --- | --- |
| SignedValue | `SEQUENCE { TBS, signature BIT STRING }` | `SEQUENCE { TBS, cert OCTET STRING, alg OID, signature BIT STRING, [0] timestamp OPTIONAL }` |
| TBS | version、seal、time、hash、property、cert、alg | version、seal、time、hash、property；可选扩展 |
| time | 上游容器使用原始 UTF-8 `yyyy-MM-dd HH:mm:ss` BIT STRING；本测试配置约定 UTC | UTC GeneralizedTime |
| 印章日期 | UTCTime | GeneralizedTime |
| 制章签名原文 | `DER SEQUENCE { sealInfo, makerCert, alg }` | `DER(sealInfo)` |

只读解析保留 V1 原始时间，不推断时区。V4 完整证书列表及摘要列表可作结构检查，
可选字段通过 `ToBeSignedDer` / `SealDer` / `Timestamp` 保留；本测试验证器拒绝
扩展与时间戳，不验证任何外部 TSA。支持范围是常见字段，不是完整 SES 规范覆盖。

SM2 固定命名曲线 `sm2p256v1`（`1.2.156.10197.1.301`），算法 OID 为
`1.2.156.10197.1.501`，SM3，显式 ID 为 ASCII `1234567812345678`。
文档摘要为 `SM3(原始 Signature.xml 字节)`，SM2 签署完整 `DER(TBS)`，由密码库
计算 ZA；签名是 DER `SEQUENCE { r, s }`，位串无 padding。密钥/nonce 使用
`SecureRandom`，不公开固定测试私钥。测试证书与印章有效区间为
2020-01-01 UTC 至 2049-12-31 23:59:59 UTC；签名时间取自唯一的正确命名空间
`SignedInfo/SignatureDateTime`，严格 UTC 秒精度，验证时重建比对。此时间仍是
签名者声明，不是可信时间证明。

## 限额与失败

- SignedValue 和 Signature.xml 各最多 1 MiB；DER 预扫描最多 32 层、4096 节点。
- 单个 DER OID 原始内容最多 128 字节，在 ASN.1 建树/十进制展开前拒绝超限输入。
- 证书最多 64 KiB，图片最多 512 KiB，签名最多 4096 字节，证书列表最多 64 项。
- 属性路径及 IA5 字段最多 1024 ASCII 字符；属性必须是安全的绝对包内 `Signature.xml` 路径。
- 拒绝 BER indefinite、非最短长度、截断/尾随内容、错误版本/字段类型和非零位串 padding；
  不解码图片，不获取网络证书/CRL/OCSP，不进行自动信任。
- `Parse` 对格式/限额失败抛 `InvalidDataException`；测试 verifier 对格式、配置或密码
  不匹配返回 `false`。Null 为调用错误；取消抛 `OperationCanceledException`。密码操作
  在调用前后检查取消，不提供中途抢占。核心 verifier 现有异常映射保持不变。

## 开发与本地消费验证

库和测试项目均进入主 solution / CI；默认打包入口从既有 validator 的 `PACKAGES`
读取原 11 个产品逐项目打包，可选库不加入默认发行批次。单独打包项目：

```sh
export DOTNET_CLI_HOME="$(mktemp -d)" DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1 DOTNET_CLI_TELEMETRY_OPTOUT=1
dotnet pack src/Ofdrw.Net.Signatures.SesInterop -c Release --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
./scripts/run-ses-interop-e2e.sh artifacts/ses-interop/e2e
```

脚本从本地包消费新扩展与本次 SDK/CLI，生成确定性布局样例的显式 Native/default
OFD，并分别添加 V1/V4 SignedValue，再导出 PDF/逐页 PNG。PNG 不代替 Preview。
功能测试与实际页面验收分别记录在 [本票证据](../evidence/ses-interop/README.md)。

固定上游为 [`5fe9c4276c64e40b455e6ea649b695adf8a9a734`](https://github.com/ofdrw/ofdrw/tree/5fe9c4276c64e40b455e6ea649b695adf8a9a734)，
结构契约见 [`SESV1Container`](https://github.com/ofdrw/ofdrw/blob/5fe9c4276c64e40b455e6ea649b695adf8a9a734/ofdrw-sign/src/main/java/org/ofdrw/sign/signContainer/SESV1Container.java)
及 [`SESV4Container`](https://github.com/ofdrw/ofdrw/blob/5fe9c4276c64e40b455e6ea649b695adf8a9a734/ofdrw-sign/src/main/java/org/ofdrw/sign/signContainer/SESV4Container.java)。
SM2 独立已知答案来自 [BouncyCastle 2.6.2 SM2SignerTest](https://github.com/bcgit/bc-csharp/blob/release-2.6.2/crypto/test/src/crypto/test/SM2SignerTest.cs)。
BouncyCastle 版本固定 `2.6.2`，MIT，实际包元数据/许可/哈希保存在证据目录。
