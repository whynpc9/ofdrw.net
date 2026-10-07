# 11 票密码契约设计

设计：gpt-6-astra high 子代理；实现：主会话 gpt-6.1-sol high。
基线 df71c7f，固定上游 5fe9c4276c64e40b455e6ea649b695adf8a9a734。

固定上游核对 UserPasswordEncryptor、UserPasswordDecryptor、KDF、OFDEncryptor、
OFDDecryptor、Encryptions/EncryptEntry/DecyptSeed/UserInfo/ProtectionCaseID。
上游实际拼写 DecyptSeed 保留；方案 1.1.1 显式写入（不复制上游 setter/遗漏缺陷），
页筛选不复制上游始终写 All 的行为；不复制缺失密文跳过或映射解密失败当明文的分支。

契约决定：net8.0 独立包 + BC2.6.2 MIT，SM3 counter KDF/SM4-CBC-PKCS7，
系统随机 FEK/IV，单 DocBody/用户/层，显式条目和零基 Page@BaseLoc 筛选；
共享资源不自动扩选，签名需先显式清理。bytes 返回完整结果，file 同目录0600暂存/rename。

用户要求损坏原子失败，CBC自身不能可靠检出中间块损坏。采用私有profile v1，
加密映射内记录全包原始文件清单/Length/SM3，用于恢复字节故障检测；严格核对输入
与恢复集合，无弱化回退。此私有扩展不声称GM/T完整性协议、MAC、认证加密、
厂商互认或主动攻击防护；攻击者整体替换/重构不在保证内。

固定向量由设计代理以OpenSSL 3.6.4独立求值；生产内部原语测试直接核对这些常量：

- SM3 abc: 66c7f0f462eeedd9d1f2d46bdc10e4e24167c4875cf2f7a2297da02b8f4ba8e0
- KDF16 12345678: 34d6effbd5bdc7c8020f619bfd8303b5
- KDF16 口令🔒: 786851358d2b990766550167a3f8981c
- SM4 key/plain 0123456789abcdeffedcba9876543210 单块: 681edf34d206965e86b3e94f536e4246
- IV 000102030405060708090a0b0c0d0e0f，包装上述FEK: 033ff79878445981b81938ec7d4f901c58885f048714e56831e5e1c1349f128c
- 同FEK/IV CBC/PKCS7 abc: 4301693c448c7da7cff13f84690f7dea

完整公开限制见 ../../password-crypto.md。此设计记录不是功能/视觉或独立实现验收。
