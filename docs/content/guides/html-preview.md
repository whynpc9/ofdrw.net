---
title: "静态 HTML 预览"
description: "将选定的 OFD 页面导出为有界、自包含、无 JavaScript 的静态 HTML，并配置页选择与资源预算。"
sidebar:
  order: 100
  label: "HTML 预览"
---

`Ofdrw.Net.Converter.Svg` 的 `OfdToHtmlConverter` 将多页 OFD 导出为静态 HTML。页面使用内嵌 SVG，不需要额外包或 JavaScript。输出是绘图层投影，不执行 OFD 动作、导航或签名信任 UI。

## 导出选定页面

```csharp
using Ofdrw.Net.Converter.Svg;
using Ofdrw.Net.Converter.Svg.Converters;

var exporter = new OfdToHtmlConverter(new OfdHtmlOptions { Title = "Report" });

await using var input = File.OpenRead("input.ofd");
await using var output = File.Create("preview.html");
await exporter.ConvertAsync(input, output, new[] { 2, 0, 2 });
```

第三个参数使用零基页索引；上例按“第 3 页、第 1 页、第 3 页”输出。传 `null` 选择全部页面，显式空列表会失败。顺序和重复页保留，每个页面出现次数都有独立的裁剪 ID 和嵌入字体族。

输入只加载一次。相同嵌入字体仍会按页面重复写入，因此大字体和重复选页可能很快耗尽预算；字体解析缓存仅在当前调用内复用。默认选择第一个 OFD 文档，多文档选择见 [Multi-DocBody](https://whynpc9.github.io/ofdrw.net/guides/multi-docbody/)。

## CLI

```sh
ofdrw ofd-to-html input.ofd preview.html --pages 3,1,3 --title "Report"
ofdrw ofd-to-html --input input.ofd --output preview.html \
  --max-pages 200 --max-output-bytes 67108864 --max-page-bytes 16777216
```

CLI 页号从 **1** 开始，支持 `1-3,2` 这样的升序范围。空片段、倒序范围、零、溢出、超长选择表达式和超预算展开会在打开输入前失败。

HTML 必须显式使用 `ofd-to-html`，通用 `convert` 不会根据 `.html` 自动路由。CLI 在目标目录暂存并原子提交；退出码为 `0`（成功）、`1`（失败）和 `130`（取消）。

## 默认预算

所有限制必须为正数：

| 限制 | 默认值 |
| --- | --- |
| 页面出现次数 | 200 |
| UTF-8 输出 | 64 MiB |
| 单个页面部分 | 16 MiB |
| 预留生成节点 | 100,000 |
| UTF-16 生成容量 | 32 Mi code unit |
| 原始嵌入资源 | 32 MiB |
| 标题 | 1024 个 UTF-16 code unit |
| 累计解码图片像素 | 100,000,000 |

节点、字符和资源预算会在渲染分配前按保守上界预留，因此可能拒绝理论上能够序列化到限制内的输出。重复页面会重复计入资源和生成成本。它们是操作预算，不是整个进程的内存、CPU 或运行时间保证。

## 安全与媒体边界

导出器生成自己的受限 SVG DOM，并校验允许的元素、属性、数值和页内裁剪引用；它不是任意 SVG 清理器。标题和文字按上下文转义，CSS 只来自固定模板及受控字体名/Base64。

输出不包含脚本、事件属性、外部 URI 加载、任意 CSS、嵌入式 SVG 图片或外来命名空间，并带防御性 CSP。未知可见绘图、不支持的字体/字形/批注语义和未知绘图扩展会明确拒绝。

内嵌图片仅接受受限的静态 PNG 或 8-bit baseline Huffman JPEG。渐进/多扫描/算术 JPEG、索引/交错/动画 PNG、私有或文本 PNG chunk、尾随数据、BMP/TIFF/GIF 等会拒绝。准入检查证明的是受支持语法，不保证浏览器解码器或色彩管理一致。

## 流写入与查看

完整结果先序列化为无 BOM 的有界私有内存暂存，再复制到调用方流。生成、选择、策略或预算失败不会改变公共输出；最终复制或 flush 失败可能留下前缀，调用方必须丢弃该输出。流保持打开并从当前位置使用，不要通过不同包装流同时读写同一底层存储。

窄屏会按比例缩放页面；打印 CSS 将每次页面出现放在独立分页中，混合纸张尺寸仍取决于浏览器和打印设置。未嵌入或不可用的字体名依赖浏览器/宿主，不构成字体身份保证。

历史浏览器与 PDF/Preview 对照只适用于其记录的实际产物和页面，见[实现验收记录](https://github.com/whynpc9/ofdrw.net/blob/8f27f92b51b70a9b01082da35ff1cfe989d46123/plans/029-followup-implementation.md)。
