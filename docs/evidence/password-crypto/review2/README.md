# Review2：首轮有效反馈修复

源码基线 dcaf640，按已收齐的 Codex/Cursor 首轮修复：

- ZIP 完整大小受 min(MaxOutputBytes, Load.MaxInputBytes) 约束，含中央目录/结束记录；成功输出可在相同预算下回读。扩展配置要求有限 MaxCompressionRatio>=1。
- 所有已返回的载入数组先登记到 OwnedBuffers，再校验任意路径；finally Dispose 清零，不清理调用方输入或已返回结果。共享loader尚未返回的缓冲及托管/后端内部内容不扩大擦除保证。
- tag 发布配置的默认 Pack 块从原 validator.PACKAGES 生成临时清单再打包，失败即停止，原精确11消费/校验/push条件保留。只运行本地 Pack 块及本地 feed 消费，未 tag/dispatch/发布。

新专项54/54、全套454/454、Python8/8。预算回归包括较小压缩输入产生较大存储输出的原bug、完整ZIP大小恰好成功/少1字节失败、文件目标原子性；数组清理回归包括unsafe前/中/后三组和正常处理中不清空；发布受控测试证明只11默认产品且生成失败不进入pack。实际原发布Pack块产生11产品，再用该feed真实独立consumer/CLI验证，通过。

可选包0.1.0-issue11.review2三nupkg经主会话与gpt-6-sol low两个独立缓存、仅Password PackageReference消费者运行通过。四组native/default×partial/all各8条目逐字节恢复、错口令/预取消保持旧输出；新增input budget=11,602,908时加密在发表前拒绝且保持sentinel。哈希/大小/nuspec依赖/readme均核对。独立路径 `/private/tmp/ofdrw-issue11-review2-independent`，结果见 independent-functional.json。

新确定性产物在 `artifacts/password-crypto/review2-current`：12张新PNG分别与新基线及此前已检查像素逐字节相同；全部新恢复OFD整ZIP字节也与此前验收的恢复OFD完全一致，原始8条目载荷没有变化。PDF重新导出，仍按要求重新打开检查。

**本轮Preview尚未完成**：已确认baseline-native/default各1,2页，4/12；四恢复PDF8页及锁定后窗口清理pending。CUA报告Mac已锁，自动解锁失败，已请求人类手动解锁。过程中出现的辅助PNG集合和直接OFD/PostScript窗口未计入PDF验收，原因未确认；按绝对root Window URL与页码重新确认，不能用像素一致或自动化替代剩余Preview。

本包只私有profile v1自闭环，无profile标准/上游/厂商包不支持；无厂商互认、非认证、非长期保存。恢复库存仅故障检测，非MAC或GM/T完整性协议。状态/代码与产物哈希见acceptance.json。旧失败与之前完整12页记录保留在父目录，不能以旧记录宣称本轮整体关闭。
