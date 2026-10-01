# Issue16 公开表格验收（2026-10-02）

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
