# Issue16 公开表格验收（2026-10-02）

**审阅入口：** [可直接查看的五份实际PDF、11张页面图及原始OFD产品档案](public-tables-evidence/README.md)。本文全部本地产物快照和哈希在该跟踪目录中按原字节保存；干净checkout可查看或提取，不依赖被忽略的`artifacts/`。

## 基线与范围

- 独立 worktree：`/Users/wanghongyi/.codex/worktrees/0ee7/ofdrw.net`；分支 `codex/public-tables`。
- 依赖 PR #9 保持开放未合并，起点准确为 `5d71dc9e4fb3ccfa4dd028386436e58cd261bc39`。本会话读取 GitHub 核对 head、六项绿色检查和 22/22 resolved review 线程；依赖视觉证据在该提交的 `flow-layout-2026-09-29.md`。
- 用户授权 stacked PR base=`codex/public-flow-layout`；不合并、不发布、不推进下一票。
- Astra High 先做源码设计与修复复核；Sol Low 独立功能与包消费验证；主任务实现及本次页面验收。
- 首版有限范围：[公开契约](../public-tables.md)。RowSpan 明确拒绝，DOCX vMerge 诊断或 Throw。Native 对不完整网格行（含未适配的 gridBefore/gridAfter）严格失败，不扩大任意 Word 表格保真。

## 功能与自动渲染

- 第一次 Layout 72/72、DOCX 130/130；修复后 Sol Low 独立全套 **295/295**（7 程序集，0 失败/跳过；Layout74、DOCX132、PDF52，其他37），日志/TRX `artifacts/public-tables/independent/full-tests-after-fix.log` 与 `independent/test-results/{Core,Packaging,Converter.Pdf,Signatures,Converter.Docx,Cli,Layout}.trx`（旧 independent.trx 不计入最终合计）。
- 11 个本地包隔离消费、manifest、CLI Native/DualLayer 通过：`independent/package-e2e-after-fix.log`、`independent/packages/`、`independent/package-output/`。新增外部 consumer 从 nupkg 使用 Table/Row/Cell，生成 3 页水平合并/底色/对齐表，核对文字完整。未发布公共 NuGet。
- 真实缺陷回归：旧多 M 路径经 PDF 导出出现斜线，独立 console 同页内容将单边路径合并后产生 **1810** 个格内黑像素；修复后 **0**。证据 `independent/negative-repro.log` 及 `negative-repro/`。PDF 像素回归另检查两格填色及预期边框，防空白产物假通过；新增正向断言最终重建后 1/1 通过，抗锯齿边框在几何坐标附近取最暗像素；随后七程序集最终 TRX 合计仍为 295/295。
- Astra 复核极薄空行越框、合并列浮点最右边越框、继承样式格内分页诊断三个 P2；修复及针对回归已闭合。不支持的格内分页从已解析格式检查，warning/Throw 均覆盖。
- SDK 10.0.401，macOS，按 AGENTS 设置 writable DOTNET_CLI_HOME、关闭首次体验/遥测、单节点及禁用 build server；NUGET_PACKAGES 显式复用现有缓存。VSTest 沙箱 SocketException 权限限制，关闭 build servers 后在沙箱外运行。初次 NuGet 网络不可达是环境失败，复用缓存后构建成功；未清全局缓存。

## 页面验收进展

样例 `e2e/Ofdrw.Net.Layout.E2E/PublicTableSamples.cs` 生成 public 表格及同文 DOCX Native/default；还重用 `generated-layout.docx` 显式 Native/default。所有 PDF 都由本次 OFD 导出；程序分别核对完整 OFD 原文与 PDF 无重复文字，没有使用直接 DOCX→PDF 替代。

首轮 `artifacts/public-tables/initial/` 生成五组共 11 页。实际在 macOS Preview 打开 `tables-public.pdf` 第一页发现斜线缺陷，不能通过；已关闭旧窗口。该轮产物和页面 PNG 保留作为失败证据。修复后必须重新生成、打开并逐页复验。

## 修复后实际验收

- 生成代码固定为 `1f846adf59c6ef6e140d184470c56a6728937e46`（生成时相同代码尚未提交，随后提交保存这一基线）；后续只有验收文档变更。样例源码、模式和限制保持以上范围。
- 本轮完整重新生成到 `artifacts/public-tables/fixed/`，逐份新开 PDF，核对 Preview 文件 URL 和 Page X of Y，再逐页正文目视；五份 **11 页**全部查看并关闭对应窗口。Cua 原始 UI/截图证据在本会话记录，机器验收记录 `fixed/acceptance.json`。
- 自动渲染：五份 OFD→PDF 全成功；公开表 415 非空白字符/3 页，Native/default 同文各415字符/2页；既有基准 Native/default 各189字符/2页，OFD与PDF原文比对一致。
- PNG：`pdftoppm -scale-to 1300 -png` 的11张 `fixed/pages/*.png` 全部目视复查。PNG仅是辅助；上述Preview为实际App验收。

| 本次 native OFD → PDF → macOS Preview | 页码 | 实际结果 |
| --- | --- | --- |
| `fixed/tables-public.ofd` → `tables-public.pdf` | 1–3 / 3 | 合并标题没有内部竖线；首/续页底色、边框闭合，01–05/06–11整行分界；三种水平/垂直对齐清楚，红粗体及斜体局部隔离，斜线已消除。第三页是正常表后正文，无空白尾页。 |
| `fixed/tables-native.ofd` → `tables-native.pdf` | 1–2 / 2 | DOCX显式Native，01–10/11整行分界，合并、底色、边框、混合中英与局部样式完整，第二页含完整第11行与表后正文。 |
| `fixed/tables-default.ofd` → `tables-default.pdf` | 1–2 / 2 | 默认Native路径与显式Native逐页外观一致；无裁切、重叠、缺字、斜线或重影。 |
| `fixed/baseline-native.ofd` → `baseline-native.pdf` | 1–2 / 2 | 本次重新转换generated-layout；第一页中英标题、斜体副标题、蓝色表头与蓝外框/灰内框；第二页局部红粗体与右对齐日期保持。 |
| `fixed/baseline-default.ofd` → `baseline-default.pdf` | 1–2 / 2 | 默认路径两页正文、表格与样式与显式Native一致。 |

以上样例不包含图片、页眉页脚或页码，不据此声称这些视觉已验收。未使用目标OFD桌面阅读器验证原始OFD互操作。功能、自动渲染、PNG视觉及Preview四项在所列范围均通过；最新PR review/CI尚待闭合。

## 产物与体积

完整清单/哈希含样例DOCX、OFD、PDF和11张PNG：`fixed/artifact-manifest.tsv`。原DOCX源 SHA-256 `17bea68d57776b02e5c68c8915de11b42087809d5310f8fafcbb003db08d7bb6`。公开表边段修复使OFD由4256增加到4565字节（+309），PDF由166859增加到166952字节（+93）；DOCX PDF保持100615/117510字节，原始OFD约15.3MB仍来自既有全量字体嵌入，无异常体积增长。

| 文件 | 字节 | SHA-256 |
| --- | ---: | --- |
| `tables-public.ofd` | 4,565 | `d164be3e7a3743834bd049c7015ffa0ede31d749525438de1bb8d8c53987aa7f` |
| `tables-public.pdf` | 166,952 | `ceaf00c958fe01b22e756e1b60254b6491d7e95c2785d0d014a4b3dfd9b6de22` |
| `tables-native.ofd` | 15,290,018 | `3e57250361f3a18cc33245a2e3e93008fb8e9599ec60dfa4a425ebcaa9a58f83` |
| `tables-native.pdf` | 100,615 | `84394c62d64c96a891ee463b60b6ab89bd32d61f79065e5cee811595c0975e24` |
| `tables-default.ofd` | 15,290,015 | `0e482a3eac8403538ba2d26b522639c6e3578c44fce8b6cbefc0cb24e98d18e8` |
| `tables-default.pdf` | 100,615 | `0fd33ec2aedfe15ec14a6856c80bc277191620d2b8ccee33b7b66a9dd0e24e9d` |
| `baseline-native.ofd` | 15,288,607 | `2570a58387aa7b5f0c0541107115aa72121c9b591502ae69bc08cda5c86ef1fb` |
| `baseline-native.pdf` | 117,510 | `322862edd1c094f81d79ee35d729979f15d24be5f1767720c31325d7c8c134a3` |
| `baseline-default.ofd` | 15,288,611 | `333ba3fad8a8514161ee7979f8048f9f6c3261b9706f22962058490f7a7a4ae4` |
| `baseline-default.pdf` | 117,510 | `4bdbe7d287b6a0332753d71c4ac823e1d0a63928e1cc43546b7594eb6e8a123c` |

## 首轮 PR review

PR #10 base=`codex/public-flow-layout`，首轮 head `5a17b853c2e14aba5e894a82fabae00dab6b93a5`。五项 CI 全绿；Codex 完成无意见。Cursor 提出三个诊断/文档问题：空单元格预算用词、Native 网格失败显示行号与具体原因、严格拒绝消息不应声称完成降级。已修正并新增三项诊断回归。绘制与测量未改变；后续仍重新生成、重新打开本次产物复验，等待新 head review/CI 闭合。

## review1 修复后最终本地复验

- 源码基线：`31f0b75fc6b2846b6651591cf006834b91ae7f4e`；改动限网格行号/具体原因、严格拒绝与降级消息分离、预算文档，未改测量或绘制。
- Sol Low 独立全套 **298/298**，0失败/跳过（Core5、Packaging23、PDF52、Signatures4、DOCX135、CLI5、Layout74）；主任务核对七份TRX。11包 `0.1.0-tables.review1` 隔离消费 E2E 通过。证据 `artifacts/public-tables/review1/independent/validation.md`、`seven-trx.log`、`test-results/`、`package-e2e.log`。
- 本轮样例重新构建0警告/错误，完整生成五组到 `review1/current/`；OFD与PDF原文比对全部通过。五份PDF重新打开、核对review1路径并在macOS Preview逐页检查 **11页**：tables-public 1–3，tables-native 1–2，tables-default 1–2，baseline-native 1–2，baseline-default 1–2。表格、样式、整行分页及基准正文保持以上预期，无新增视觉缺陷。对应窗口和残余打开面板已关闭，Cua确认 noWindowsAvailable。
- 本轮11张PNG与实际目视过的 `fixed/pages/` **11/11逐字节一致**，为自动渲染辅助证据；本轮Preview为实际新文件逐页复验，没有拿像素一致替代Preview。
- 最新清单/哈希与环境、页码：`review1/current/artifact-manifest.tsv`、`acceptance.json`。PDF体积166952/100615/117510字节保持，Native OFD因ZIP时间戳仅有个位数字节压缩差异，无异常增长。

| review1/current 文件 | 字节 | SHA-256 |
| --- | ---: | --- |
| `tables-public.ofd` | 4,565 | `f691f2ae8536e9e4c6de0d6d2cd81d63ac4849c1b24fc9c3e02284c5c682a322` |
| `tables-public.pdf` | 166,952 | `187176a8d12830d088fba63b72e17dcb0f687bb1a365de20011bdc6bb266861c` |
| `tables-native.ofd` | 15,290,019 | `ca3e335517ade22049eafdd2d2e95931a17c40afbce11b656c8c87608ecae56a` |
| `tables-native.pdf` | 100,615 | `160fed4074d198e2490b1823dcb023c25ba81427715fdbd8c04524fcc528dda9` |
| `tables-default.ofd` | 15,290,015 | `831137df57123cd65d23e8f74945d56716d68cd3d83f9914e8bc4c43e59b725a` |
| `tables-default.pdf` | 100,615 | `c3c6f14942cfc7dbbf4238c19e56cc8e477da1c9f87b6c44da730e4c0c2ae00e` |
| `baseline-native.ofd` | 15,288,606 | `26b303c596efe5883ab3f8f94fc90a5ce2abee7ca4fa8ef8f2e2ac12b7c5d4ad` |
| `baseline-native.pdf` | 117,510 | `16d06f1b9692408bdc3ae473836ec585dc70505c476d9729d6bce4526c2198ee` |
| `baseline-default.ofd` | 15,288,610 | `03097d7540edff62caccae7ae01ca62ae9ce4ad5c9e572848035619fd1ee38ac` |
| `baseline-default.pdf` | 117,510 | `cc6d6827fbf0cffb19af5612565c0ab7c2a625be45d9ad2e94e1ff2c6a92b62f` |

## review2 最终诊断复验

- 最终生产源码基线：`26f33e91d716c863bed09e5feb3c95a0a423b00f`；后续仅文档。Cursor复审指出非正gridSpan在Reader提前失败，仍需行/格位置；已修复，0/-1第二行定位和旧span0定位断言通过，格内分页降级测试也核对实际警告文本。
- Sol Low独立全套 **300/300**（Core5、Packaging23、PDF52、Signatures4、DOCX137、CLI5、Layout74），0失败/跳过；主任务核对七份TRX。11包`0.1.0-tables.review2`隔离消费、manifest/bytes、CLI和Native/DualLayer通过。证据`artifacts/public-tables/review2/independent/validation.md`、`full-tests.log`、`seven-trx.log`、`test-results/`、`package-e2e.log`。
- `31f0b75→26f33e9`生产代码只对非正gridSpan错误增加行/格位置，格式、测量、分页、绘制以及有效文档路径未变，Sol独立静态核对一致。没有受影响的正常页面。本轮没有重新生成/打开新Preview；**实际最新Preview范围仍为31f0b75生成的review1/current五份11页**，不冒称26f33e9再次查看。
- 本票首轮三个意见和复审非正跨度意见均已修复、分别有回归；最新交付head的CI/线程/复审以[PR #10](https://github.com/whynpc9/ofdrw.net/pull/10)的读回为准，保持开放未合并。完整记录包含功能、自动渲染、PNG辅助与实际Preview各自范围。

## review3 证据可访问性修复

Codex指出忽略目录中的产物无法从干净checkout复核。现将本会话实际文件按原字节保存在跟踪的 `docs/validation/public-tables-evidence/`：五份最新PDF、11张PNG和页码/哈希记录可直接查看；`products.tar.zst`含initial/fixed/review1/current原始DOCX/OFD/PDF/PNG及各次独立日志、最终七份300测试TRX和11包清单。长窗口压缩仅去除重复字节，未重写产物，展开约189MB、档案约16MB。

`bundle-manifest.json`列出108原始文件的大小/SHA-256及档案自身哈希，`scripts/verify-public-table-evidence.py`已验证108原始文件和18直接副本，提取后OFD/PDF的原始哈希保持。初次打包出现macOS自动AppleDouble元数据，被校验器拒绝；以COPYFILE_DISABLE=1重建档案后完整通过，未放宽校验。CI增加zstd及档案完整性门。此轮仅证据、校验脚本和CI，不改生产转换/排版源码；实际Preview范围及300功能回归仍为上述基线。

Sol Low对`8a2923c`从Git仅导出校验脚本和跟踪证据到干净临时目录，独立验证108+18文件；提取五组OFD/PDF与原产物原始哈希逐字节一致，七份最终TRX合计300/300。损坏PDF、缺失PNG、篡改档案和不安全路径四个负向均exit1拒绝。可复查记录/日志也已跟踪：[独立证据验证](public-tables-evidence/integrity-validation/validation.md)。未改变或重新生成页面，本轮无新的Preview检查。
