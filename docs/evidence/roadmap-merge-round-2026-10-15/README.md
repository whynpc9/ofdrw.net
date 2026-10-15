> 本归档保留排除前运行源码 17dea14 的历史验收，不作为排除版当前全套计数。人类已选择本轮排除 Issue11；当前分支无口令包源码或其接受 head 祖先。其余产品源码、字体、页面字节未变，Preview 原基线保留。排除版重测、祖先及范围差异见 [exclusion-readback.json](exclusion-readback.json)。

# 首轮集成验收证据

运行源码：`17dea14282ee58c00051e36a3257449102883d7b`。完整范围、限制和每份产物的实际基线见 [acceptance.json](acceptance.json)，环境见 [environment.json](environment.json)。这是合并前证据，不是 NuGet 发布记录。

- [evidence.tar.zst](evidence.tar.zst)：111 个可审阅载荷，约 22.8 MiB。
- [manifest.json](manifest.json)：每个载荷的相对路径、体积、SHA-256。
- [archive.json](archive.json) 和 [readback.json](readback.json)：归档身份与逐项解包读回结果。
- 2,099 项完整回归通过，12 项可选未执行；另强制字体 65/65、零跳过，与完整回归重叠。
- 实际默认 11、SES 12、vector 13 包消费。样例输出、包依赖和原始失败/最终成功日志都在归档中。
- Preview：33 页样例范围全部逐页检查并关闭窗口。其中不受最终空格校验改动影响的 graphics/typed 记录保留 `94a2d9a` 基线；最新 vector 15 页与空格 2 页基于 `17dea14` 重新生成、重开。不得把不同基线描述成同一次生成。
- PNG 自动渲染完整；PNG 独立目视仅 graphics 10 页、空格 2 页。其他样例实际 Preview 验收通过，不把 PNG 或文本提取当 Preview。
- fallback 第 1 页的裁剪仅保留页脚，第 4 页源内容为空，均是明确的边界样例；第 2 页 40% 透明文字、第 3 页实际图像及第 5 页原生文字按预期出现。
- 空格修复样例字体是仓库原创的矩形测试符号，A/B 显示为两个黑色块是预期；窄 NBSP 槽位不得出现第三个缺字轮廓，普通/粗体均与分开 A/B 基准像素一致。
- 公共样例的字体仅 Noto OFL 或仓库 MIT 原创字体；完整许可和生成来源在归档 `licenses/`。本地 original-v1 诊断仅保留页面图/哈希，不包含原 OFD/PDF/SVG 二进制。

解包：

```sh
tar --zstd -xf evidence.tar.zst -C /your/empty/output-directory
```

归档使用 Zstandard long window 27 压缩重复的公开字体载荷；解码需要支持 Zstandard 的 tar/zstd。逐项身份以 manifest 为准。功能结果、渲染检查、PNG 目视和 Preview 在 acceptance 中分别列出；范围仅限这些确定性样例。正式发布记录仍为 `not-reviewed`，没有发布 NuGet、创建 tag 或触发 dispatch。

## 排除版当前结果

[PR22](https://github.com/whynpc9/ofdrw.net/pull/22)，运行源码 0a21fbe：2,033 全套通过/12 可选未执行、强制真实字体 65/65 零跳过、20 脚本用例、默认 11/SES 12/vector 13 真实包消费通过。只有独立口令项目被排除，其余产品与已有实际 Preview 基线源码逐文件相同；新 15 页 PNG 读回一致。原 33 页实际 Preview 记录保留其原始源码/字体/页码，本次不虚报重新查看 GUI。

[排除版读回](exclusion-readback.json) 列出 11 个接纳 head 和两个明确排除 head、全部范围差异、原始 TRX 计数及 [32 文件验证归档](exclusion-validation.tar.zst) 的逐文件哈希。旧 evidence.tar.zst 不改写为新候选全套结果。最新远端 CI/review 仍须在 PR22 单独读回。

## 最新 SVG 半透明空白修复

运行源码 8018d70：2,039 全套通过/12 可选未执行，六项外部 OFD 空白 alpha 回归修复前失败、修复后通过。新 0.1.0-merge.20261015.5 默认 11/SES 12/vector 13 包消费通过。当前 SVG PNG 已目视，对应当前 OFD→PDF 一页在 Preview 查看并关闭；加上原明确基线的 33 页，选定范围共 34 页。[完整读回](svg-alpha-readback.json) 与 [原始日志/TRX/当前页面/包清单](svg-alpha-validation.tar.zst) 逐项解包校验。首次排除报告和 17dea14 归档保留其历史身份，不重写成当前结果。

## tag 工作流的强制字体入口

[同一 runner 的实际 65/65 零跳过结果及顺序读回](release-font-gate/readback.json)。工作流在候选版本 Build/Pack 之前运行门禁，默认失败阻止后续发布；always-upload 保留记录。只改工作流和证据，产品运行源码仍 8018d70。没有运行 tag、dispatch 或发布。
