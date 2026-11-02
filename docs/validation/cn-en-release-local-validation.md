# 中英文候选本地独立验证（2026-11-02）

> 历史验收记录：以下“未达到可发布验收”及 FAIL/未完成结论属于当时的 34 页矩阵和 `1e6f5eef` 集成输入，原样保留。2026-11-02 用户另选限定物理 Regular/Bold 的新候选范围，见[当前计划](../next-steps-2026-11-02.md)；这不是重新判定旧运行通过，也不是 preview.10 的新构建或视觉验收凭证。

**结论：未达到可发布验收。** 当前固定样例的原 Regular 表头“量”字在 144 DPI 页面中仍有字腔堵塞；新生成 native/default PDF 的页末日期在 macOS Preview 中仍无法得到可操作的查找结果。原始剪贴板字节未能验证。以下结果只覆盖指定输入和本机，不是新 main、三平台或公开发布证明。

验证者为 `gpt-6.1-terra`、High；macOS 26.6.2 (25G83)、.NET SDK 10.0.401，GUI 由该验证者独占。集成运行源码 `1e6f5eefb60b18ea9951769273b749a27ac9c174`；全套测试运行于 `61120ae19f0e57b2280f9222c305b512eb64de5c`，两者差异仅 N13 E2E 颜色比较与其文档，`src/`、测试及样例未变。运行后本报告不属于上述源码身份。产物在忽略目录 `artifacts/cn-en-release/local-validation/`，长期保存需依仓库存储规则另行处理。

## 按任务核验

- **N00**：独立核对 F01（3,347 字节，SHA-256 `17bea68d57776b02e5c68c8915de11b42087809d5310f8fafcbb003db08d7bb6`）、可信字体 ZIP、两字体和两许可证的长度及哈希，均符合 N00。物理 Bold manifest SHA-256 `23da8dc97db7c6b73c2f71fc9da15ed1e18fe0701eb498ea5152eab26ad8a3b8`，五个文件哈希独立核对；其生成过程和字体元数据未独立重建。矩阵为 34 页。
- **N01**：部分静态复核确认所列区间 19 commits、41 文件、+3917/−82，两个 main blob 与报告相符，SVG 的 family 初始化与报告一致；未独立逐项复核全部处置或动态验证 C1/C2。
- **N03**：独立核对生产者保留的 12 个 OFD 的大小、哈希、XML 中正字号及 TextCode；4 个拒绝案例未单独重跑。4 个普通 4 mm OFD 分别导出 PDF、144 DPI PNG 并在 Preview 检查，`AV` / `To` 可见无明显裁切重叠；0.0005/0.001 mm 仅判序列化，不以可读性裁判。产物 `n03/`，输入身份见 N03 manifest。完整测试包含 N03 回归。
- **N02**：基准 `b02162be` 的 native/default PDF 在 Preview 中，`2026-01-01` 和 `待复核` 同时显示 Found on 1 page 与 Not found，导航禁用且无高亮；`Alpha`、`测试日期`、`2026` 可高亮；`01` 仅定位第一个；负控 `2026-01-02` 无结果。记录 `n02-baseline/preview-find.json` SHA-256 `5d1d2bae2eacb64e058c7ec64c62726b0cfa03727b441b4bb9493e1976e1050b`。本次 `core-v2` 新生成的 original native/default PDF 中完整日期仍复现相同 UI 失败；根因未确定。
- **N13**：首次 `61120ae` 的 core 运行在颜色断言处失败，日志 `core-run.log`、首个 OFD 保留；修正 E2E 比较后于 `1e6f5eef` 重新构建运行，10 份 PDF 共 26 页的 OFD 原文、PDF 和样例断言完成。`core-v2/cn-en-release-manifest.json` SHA-256 `b58358897e9a9c690e34027fc6f5d0fb370d88e40b1d1b451a3aeed9907f4564`，记录源 DOCX、各模式 OFD/PDF、页数、字体及哈希。实际配置文件的 Regular、Bold 和 Latin 哈希与 N00 清单一致；嵌入字体子集身份见 manifest，不能与源字体哈希混同。
- **N12 本地范围**：`python3 -m unittest discover -s scripts/tests -p 'test_release_gates.py' -v` 独立 10/10 通过。新空 feed 运行 `python3 scripts/run-converter-package-e2e.py 0.1.0-cn-en-local.2 --packages-dir <feed-v2> --output-dir <package-v2>`，32 步 PASS。12 个包重新哈希与 `feed-v2/package-manifest.json` 一致，其 SHA-256 `28b1b2718d6f6d4fc7be2f29aedd45d830feaa07a3d0a74bdd39fbace6672e99`；前后 manifest 验证由驱动执行。隔离消费者为 .NET 10，receipt 记录选中 `netstandard2.1` DLL 与包内、缓存及输出的身份；HarfBuzz 版本由运行收据报告，实际加载 native 模块路径/哈希未独立核实。该本地版本非公开发布。

## 功能与页面

按仓库规则设可写 `DOTNET_CLI_HOME`、遥测和首次体验环境、显式现有 NuGet 缓存，使用单节点与禁用构建服务器参数。`dotnet test Ofdrw.Net.sln -c Release --logger trx ... --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false`：**4,612 通过、17 跳过、0 失败**；`scripts/run-trusted-font-tests.py --no-build` 单独 **74/74 通过、零跳过**。日志和 TRX 保留于 `full-tests.log`、`test-results/`、`trusted-tests.log`、`trusted-tests/`。还原报告 SixLabors.ImageSharp 2.1.13 的 NU1902/NU1903 漏洞警告，不能由测试通过消除。

准确集成 SHA 上执行 `--cn-en-release-only`、现有 Flow 默认生成器及 `--flow-fixed-only`；输入、字体、OFD/PDF 与程序断言见各输出和日志。**34/34 页面已由 `pdftoppm -png -r 144` 渲染并逐页 PNG 目视**；`png/manifest.json` SHA-256 `e44d075a6d592448d0e183fb01ad297c9307d1335d753b376d1b9335fd758ca2`。原 Regular 的 native/default 第 1 页“量”仍 FAIL；单独物理 Bold 的两模式该字清晰，不能覆盖原配置失败。其余已检查的 Flow、中英文局部样式、对齐、双色图、表格、页眉/页脚/页码及选择页顺序未见明显缺陷。

**真实 macOS Preview 检查 34/34 页**：`core-v2` 的 original/physical layout 各 native/default 两页、headers 双模式各五页、image 双模式各一页、selected 双模式各三页；另 `flow-full/flow-public.pdf` 三页、三份 alignment 各一页及 `flow-fixed/flow-fixed.pdf` 两页。DOCX 查看链路为本次 DOCX → native/default OFD → PDF → Preview；Flow 为 OFD → PDF → Preview。按页查看正文和所涉样式、图片、表格、页眉页脚/页码、页序；截图在操作记录中，PNG 另存于上述目录。原配置的 144 DPI 量字缺陷使整体视觉仍 FAIL，检查页数不等于视觉通过。最终任务窗口关闭，打开面板退出后 Preview 进程已不可用。

**Copy/Find 单列**：新 original native/default 的日期 Find 失败如上。Physical native/default 实际全选复制后粘贴到 Preview 搜索字段，字段中英文标题出现字母间空格；搜索字段可能转换文本，不能作为原始剪贴板字节或新产品缺陷的证明。当前 computer-use API 无原始剪贴板读取，工具明确禁止以其他技术操作电脑，因此原始 Unicode Copy 验收 **未完成**。未用 AX、PDFKit 或文本提取替代。未验证其他阅读器的 Copy/Find。

这些结果不填入旧 `release-preview-acceptance.json`，不宣称候选 accepted、三平台 CI 全部完成、已推包或发布。跨平台、真实加载 native 哈希、原始 Copy 和所有未完成发布门由协调者另行判定。
