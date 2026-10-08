# 19 受限 Skia 事件适配验收证据

功能、真实本地包消费及 **macOS Preview14/14页全部通过**。2026-10-08 首次连接因 Mac 锁屏失败，历史 0/14 状态保留；协调方后续重新分配独占时段后，主代理以准确 d3c6/r3 文件 URL 逐页实看八份PDF。8个本票窗口全部关闭，其他任务旧窗口保留，GUI已释放。只对这14页及列出的确定性样例作结论。

源码：04 起点 `ab87bca`，19 首轮运行时代码 `8bf97cf`，独立消费修订 `8084147`；当前运行时 `bdff0d0` 修复自定义 Blender，`e6e11ee` 的取消/cleanup修复保留。设计由 GPT-6 Astra High 完成，实施主代理 GPT-6.1 Sol High，独立验证 GPT-6 Sol Low 从 `git archive` 固定提交执行。本实现是**合作 producer 显式事件入口**，既有 `Render(SKCanvas)` 必须改造；没有透明捕捉任意 canvas/picture/PDF 的承诺。见 [设计](../../skia-adapter-design-contract.md) 与 [教程](../../tutorials/17-skia-event-adapter.md)。

- 全套 449/449，含 19 的 20 个针对性用例；Python validator 5/5。失败原子性、错误同名字体载荷、唯一 ID/face 样式、状态/超限/取消/不支持效果、快照 native 生命周期均有断言。
- Primary 本地版本 `0.1.0-skia.20261008.r4` 默认 11 包和可选 12 包 feed 消费通过。Low 在冻结源码的独立副本，以 `0.1.0-skia.20261008.low4` 重打包并执行两个干净缓存消费通过，assets 中没有 project reference。默认产品 nuspec 不含 Skia，仅可选产品直接依赖 SkiaSharp。
- 合作 producer 样例 2 页、17 PathObject、27 TextObject、0 ImageObject；原始 Unicode（包括连续空格）保留，唯一字体 ID 与实际 face 字节 SHA256 核对。适配与直接 04 的六个 ZIP 条目解压内容逐字节相同；Low 的外层 ZIP 哈希因时间戳不同，所以不声称所有重生 OFD 二进制相同。
- 本次 PDF→PNG 的适配/直接 04 像素相同；另有真正 SKCanvas 输出两页 PNG。主代理与 Low 已查看 19 的两页 PNG；局部文本/图形坐标另有像素探针。这是辅助受限样例自动/PNG证据；另有真正 Preview14页记录。
- 首轮 primary r1 / Low low1 的干净消费编译失败日志完整保留。原因是样例复制到仓库外后缺少 Directory.Build.props 的 ImplicitUsings；修订在 csproj 显式声明，修复后的真实包消费另有日志，历史未覆盖。

[acceptance.json](acceptance.json) 分开记录功能、原生对象、Preview 和 review 状态；[manifest.json](manifest.json) 绑定全部载荷的大小/哈希及 [公开审阅归档](evidence.tar.zst)。解压：

```sh
mkdir -p artifacts/skia/public-review
zstd -dc docs/evidence/skia/evidence.tar.zst | tar -xf - -C artifacts/skia/public-review
```

归档包含当前 OFD/PDF/PNG/SVG/TXT、原始失败捕捉探针、Native/default DOCX 对照、source/package/Low 清单及日志、字体与 Skia 原生许可。`candidate/` 是本次本地 nupkg 消费所得适配/直接04产物；`regression/` 是本次重生04与 licensed DOCX Native/default 对照。

已实际 Preview 的当前路径如下，共8份PDF/14页：

| 目录（本 worktree artifacts/skia 下） | 文件 | 已查看页 |
| --- | --- | --- |
| package-r4/output | adapted.pdf、direct04.pdf | 各 1–2 |
| graphics-regression-r3 | graphics.pdf、graphics-roundtrip.pdf | 各 1–2 |
| graphics-regression-r3 | graphics-name-only.pdf、graphics-fractional-clip.pdf | 各 1 |
| graphics-regression-r3 | baseline-native.pdf、baseline-default.pdf | 各 1–2 |

链路必须是合作 producer→native OFD→PDF→macOS Preview、以及 licensed DOCX→显式 Native/default OFD→PDF→Preview。后两份含中英文、局部粗斜体/颜色、填色边框、对齐和两页分页。不会用直接 DOCX→PDF 代替 native OFD 的验收。

本次环境 macOS 26.6.2、.NET SDK 10.0.401（global.json 10.0.100/latestFeature）、SkiaSharp 3.119.1、poppler pdftoppm 144 DPI 的19页图；04对照 110 DPI PDF/720px SVG。使用固定静态 OFL Noto Sans CJK SC，SHA256 `3012a9b63f5eca3e3b38f23a1be5ed504675e394abf8e7a4fa981506582c04aa`，由既有 pinned `f8d157532fbfaeda587e826d4cd5b21a49186f7c` Noto variable 源的 Regular 实例生成，公开归档附原许可。主 19 OFD 约 11.60 MB，PDF 约 79.6 KB，主要体积来自整份嵌入字体；19 不实现字体子集，体积不比直接04减少。

共享 pack 修改只移植已完整审查 PR15 `f7cbbcf991f33b6bc9c90cc118c636dbcca78404` 的两个文件中的默认产品隔离逻辑，直接复用原 validator.PACKAGES 清单。没有 merge/cherry-pick 11/12 功能分支，没有修改04绘图 API或默认 Layout/Converter 的 Skia 依赖。

复现从根目录执行（遵守 AGENTS 环境变量/单节点 flags）：

```sh
scripts/run-converter-package-e2e.sh 0.1.0-skia.20261008.r4
scripts/run-skia-package-e2e.sh 0.1.0-skia.20261008.r4 \
  artifacts/package-e2e/0.1.0-skia.20261008.r4/packages artifacts/skia/package-r4
scripts/run-graphics-e2e.sh artifacts/skia/graphics-regression-r3 # unchanged04/DOCX regression snapshot
python3 scripts/collect-skia-evidence.py bdff0d06bf9de0ce5aeca1d6da5a82b0b8042dfd 0.1.0-skia.20261008.r4 --round r4 --independent-round low4 --regression-round r3
```

需先按 `scripts/install-ci-fonts.py --directory artifacts/graphics-fonts` 准备同许可字体；独立 Low 证据的原始副本不属于通用重生成脚本。当前仅对列出的实际样例/页做判定；clip/layer/image/shaping及任意PDF事件来源仍不支持/未证明，05 与21仍保持独立闸门。未合并、未 tag、未 dispatch、未发布 NuGet，`production_release_accepted=false`。

首轮 PR17 `158a63c`：Codex 一条 P2（取消前不要推进producer），Cursor 无发现，六 checks 通过。当前 `e6e11ee` 已修，17用例与 Low 实际 nupkg observer 证实预取消 Get/Move=0、Get 中取消 Move=0、Dispose 异常零提交；旧 r2 的 Get=1/IOException 原始复现保留。当前 r3 的适配/04 PNG 与之前逐像素一致；另已实际完成 Preview14页。44635c9 的六检查与 Cursor 复审通过，Codex仅要求补Preview，现以当前实际验收履行；后续文档head仍需复核。

R4：4efc19b最新双bot均指出custom Blender getter伪装SrcOver；真实r3nupkg/nativegreenpixel复现后，bdff0d0只接受defaultnull或SDKcanonicalSrcOver，自定义runtime/arithmetic在所有primitive入口拒绝。主/Low449/20及r4/low4真实11+12/package observers通过。受影响r4适配/04两PDF四页已实际Preview重开通过，2窗口全部关闭并释放GUI；04/DOCX未改十页保留原r3验收和相同PDFhash。总覆盖14页以4新+10未变明确区分，旧14页实看与所有失败史保留。后续latestheadreview读回见最终PR闭环评论，不以本文件自称含有自身commit。
