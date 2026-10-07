# Review3：原始ZIP与超限XML所有权

基线8ab862c，Codex复审新增两项有效P2按Astra High边界方案修复：
一次有界内存snapshot从调用方当前Position读到EOF，原始ZIP FullName预检在任何
规范化/解压前发生，再同snapshot重置Position=0交现有Loader；不改Core。
反斜杠、空白文件basename、大小写重复、非规范路径、非零payload目录被拒绝。
目录占位只允许单尾斜杠且零载荷；目录不保留是既有明确契约。

复制缓冲finally清零；MemoryStream实际扩容由base先成功复制，新旧引用不同才
清零旧buffer，同Capacity/失败setter不破坏数据；Dispose幂等清当前整个capacity。
调用方stream保持打开且其backing array不清空。XmlBytes先登记owner再判断metadata
预算，3个生产调用点去重。预算并不限制序列化过程的峰值分配，不扩大为Loader
未返回缓冲或BCL/后端/immutable string不可恢复擦除保证。

66/66专项（含raw names、目录payload、当前位置/seek/nonseek、超限1字节、取消、
实际resize旧buffer、同容量/失败setter、返回copy、超限XML清理与原子文件失败）通过；
全套466/466、Python8/8通过。真实review3三nupkg经主会话与Sol Low两个独立
cache/source mapping、仅Password PackageReference消费6报告组通过：四组各8entry
精确恢复、same inputbudget加密拒绝、raw ZIP三种拒绝与sentinel保持；无暂存残留。
独立路径 `/private/tmp/ofdrw-issue11-review3-independent`；hash/deps/readme逐项一致。

自动optional tag发布P1经协调确认为范围外：票要求独立可pack并本地消费，不要求
自动加入默认tag发布。计划M5可选发布需单独判断、用户保留默认11 feed且本票不发布。
公开README/docs已准确区分源码/本地feed与未发布NuGet，不用“当前不执行”掩盖未来
默认配置；现有默认11发布Pack/消费已修好并实测，未来optional发布另行授权/闸门。

R3最终实际包产物 `artifacts/password-crypto/review3-current`：四份新恢复OFD整ZIP
字节与review2/初始已审产物相同；从六份明文OFD以未修改CLI导出6PDF/12PNG，
每页新PNG与baseline/先前像素一致。批量导出时误包含加密OFD，核心按既有false能力
拒绝缺OFD.xml/缺所选Content.xml；这些失败不计视觉结果，最终只6份明文PDF入证据。

**R3 Preview 0/12，未完成**：Mac锁定且自动解锁不可用，等待人类manual unlock；
不沿用review2已查4页冒充R3，不用PNG/字节相同冒充Preview。锁定时任务窗口清理也
pending。源/产物hash、环境/基线/许可、功能与pending范围见acceptance.json。
历史initial/review1/review2全部独立保留；不merge/tag/dispatch/NuGet发布。
