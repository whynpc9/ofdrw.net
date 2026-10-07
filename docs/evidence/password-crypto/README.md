# 11 票：私有口令 profile 验收证据

**最新状态：[review3](review3/README.md)**。raw ZIP预检/有界输入snapshot及超限XML ownership修复通过66专项/466全套/8 Python、实际review3包主/Low独立消费6组；2026-10-08在 macOS Preview 逐份重新打开最新6PDF并检查12/12页，未发现视觉缺陷，已关闭本票窗口并释放GUI。历史review2的4/12与原R3锁定快照保留，最新结论见review3外置follow-up。

基线 `df71c7f`，独立分支 `codex/gmt0099-password-crypto`，PR base 为 `codex/document-tools`。
本票仅交付 `Ofdrw.Net.Crypto.Password` net8.0 可选包，使用 BouncyCastle.Cryptography
2.6.2 MIT；固定上游 `5fe9c4276c64e40b455e6ea649b695adf8a9a734`。

本包仅支持带恢复库存的私有 profile v1 自闭环，明确拒绝无此 profile 的标准、
上游及厂商包。不是通用 GM/T 0099 支持，未认证、无厂商互认、非长期保存格式。
库存 SM3/Length 仅故障检测，不是 MAC、认证加密或 GM/T 完整性协议；不保证
检测 ZIP 元数据或库存外描述字段的逐位变化，也不保证抵御主动整体替换/重构。

## 功能验证

- Astra High 完成算法/契约设计并二次复核限制，见 [design.md](design.md)。
- Sol Low 独立仅读源码审查未发现阻断实现缺陷；随后又以仅Password PackageReference、
  独立目录/cache/source mapping实际运行四组恢复与错误口令/取消原子验证，通过。
  未使用project引用或旧Ofdrw缓存，也未代替主会话Preview验收。
- 新包固定 SM3/KDF/SM4/CBC/包装向量、部分页/全条目、损坏与原子失败专项：48/48。
- 整套回归：448/448（Core5、Packaging250、PDF79、Signatures4、DOCX49、CLI13、Password48）。
- 本次 native/default × partial/all 共四组，每份原始8条目均按名称/字节精确恢复，正文完整。
- 实际本地 NuGet 消费：Core/Packaging/Password 三包本次打包，独立临时目录/缓存，
  唯一 PackageReference 为 Password；四组恢复均逐条字节相等，错误口令与预取消保持旧文件、无暂存残留。
- 默认 Converter/Core/Packaging/CLI 的项目引用无可选密码包或 BC 后端。默认11包验收脚本
  改为从既有 validator 产品清单逐个 pack，防止加入 sln 的本票可选产品混入原11包 feed。
  原精确清单/缺包失败语义保持；此共享脚本修改由协调明确分配给11。
  本次默认11包实际打包/隔离消费与CLI安装完成，版本0.1.0-issue11.default；Python validator 5/5，包括缺包仍失败。

环境：macOS、.NET SDK10.0.401、Poppler，精确信息与源码哈希见 [acceptance.json](acceptance.json)。
Dotnet 使用可写 CLI_HOME、既有显式 NuGet 缓存、单节点/no reuse/no shared compilation。
沙箱 VSTest SocketException(13) 属 IPC 限制，关闭 build servers 后使用同旗标在沙箱外运行测试。
未清理全局缓存。

历史失败保留在测试日志：首次 surrogate InlineData 经属性编码变为替换字符，旧记录48/49；
改为运行时构造无效代理项后48/48。
最终KDF计数器2测试手工抄写oracle多1个nibble（33hex），按固定输入与大端计数器公式
由OpenSSL完整摘要截取16B、Python hashlib和BC生产路径核对为32hex；详见
kdf-oracle-correction.json。保留00:02失败TRX，00:03密码专项48/48后重跑全套。首次含21MB字体样例被默认16MiB单条目预算拒绝，
样例显式提高到单条目32MiB、总展开/输入128MiB、输出160MiB后通过；未提高产品默认预算。

## 实际页面验收

查看根路径 `/Users/wanghongyi/.codex/worktrees/587a/ofdrw.net/artifacts/password-crypto/current`。
2026-10-07 亚洲/上海时间，获协调 GUI 独占分配后逐个打开、逐页查看并关闭：

| 文件 | 已检查页 | 查看链路 |
| --- | --- | --- |
| baseline-native.pdf | 1,2 | DOCX→native OFD→PDF→Preview |
| baseline-default.pdf | 1,2 | DOCX→default OFD→PDF→Preview |
| native-partial-restored.pdf | 1,2 | DOCX→native OFD→部分页加密→解密OFD→PDF→Preview |
| native-all-restored.pdf | 1,2 | DOCX→native OFD→全部条目加密→解密OFD→PDF→Preview |
| default-partial-restored.pdf | 1,2 | DOCX→default OFD→部分页加密→解密OFD→PDF→Preview |
| default-all-restored.pdf | 1,2 | DOCX→default OFD→全部条目加密→解密OFD→PDF→Preview |

共12页：中文完整、比例英文间距正常、灰色斜体副标题与蓝色标题、表格浅蓝底色/
灰内框蓝外框、第二页局部红色粗体与右对齐日期、两页分页与基线一致；无裁切、重叠、
重影或异常空白页。每份恢复PDF的两张96DPI PNG与对应基线PNG逐字节相同。
使用本次新生成的实际恢复OFD导出PDF，没有直接DOCX→PDF代替。文件哈希绑定在
acceptance.json / manifest.json。所有本任务文档窗口已关闭，GUI时段已释放。

样例是无隐私 generated-layout.docx 的Noto字体替换版；固定字体哈希
`3012a9b63f5eca3e3b38f23a1be5ed504675e394abf8e7a4fa981506582c04aa`，OFL许可随产物。
本样例未覆盖图片、页眉页脚和页码；不推断复杂Word保真、页资源保密闭包或外部阅读器互认。

体积：原压缩OFD 11,602,908字节，恢复OFD 21,682,785，部分加密21,685,015、
全条目加密21,685,759。约1.87倍增长主要来自保守的不压缩ZIP输出；PDF各117,058字节。
仅载荷字节保留，不承诺ZIP压缩字节、时间戳/目录项/权限往返。

## 持久工件与闭环状态

`evidence.tar.zst` 保存上述OFD/密文/PDF/12张PNG/原文/许可/功能记录；manifest对每文件
记录尺寸/SHA256。解包后按manifest逐项核对；无需密码厂商。测试用口令见fixture源码，
所有内容均为确定性虚构数据。实际包消费证据单独保存在 `package-functional.json` /
`package-manifest.json`，运行日志与TRX在证据包中。

本地功能与所列页面通过不等于PR闭环。Codex/Cursor首轮、最新head复审/CI和全部有效
线程处理状态会在review记录中补齐；不合并、不发布NuGet。

## 首head后修补（PR #15）

0d88008首head保留上述功能/页面证据。协调发现共享脚本process substitution内清单
生成异常可能被循环吞掉并接受旧feed；最终先生成TASK_DIR/default-products.txt再循环，
与12票提供的共享窄patch逐字节一致。注入完整旧11包+导入生成失败：旧脚本确实red、
新脚本在pack/validator/消费前停止；Python6/6。再次实际原11门版本
0.1.0-issue11.default2通过（非只模拟pack）。

Windows首轮专项因目录覆盖报UnauthorizedAccessException而测试只接受IOException失败；
按照明确平台异常调整断言，保留暂存清理，未修改密码runtime。首次远程失败日志、旧脚本
red、修补后green与default2实际消费存于review1-validation.tar.zst及review1-validation.json。
初始大证据包不重写，所有已Preview产物逐项尺寸/hash再次核对一致，不复用重新生成页面。
首轮Codex/Cursor、最新head CI/re-review仍待关闭，不据本地通过宣称PR闭环。
