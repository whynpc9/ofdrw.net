---
title: "纯文本生成原生 OFD"
description: "使用 PlainTextOfdConverter 或 text-to-ofd 将字符串和严格 UTF-8 文本分页生成 OFD。"
sidebar:
  order: 80
  label: "纯文本生成 OFD"
---

`Ofdrw.Net.Layout` 的 `PlainTextOfdConverter` 是 `FlowDocument` 的轻量适配器，适合通知、日志和简单纯文本。它返回新的内存 OFD 包，再交给 `OfdPackageWriter` 保存。

> **可选最终文字排版（自 0.3.0 起）：** 设置 `PlainTextOfdOptions.Layout.TextEngine` 后，纯文本改用段落排版引擎，以真实字体塑形并嵌入字体子集（支持制表位、合成粗斜体、CFF/OTF 与增补汉字）。默认行为不变，见[启用最终文字排版引擎](https://whynpc9.github.io/ofdrw.net/guides/final-text-layout/)。

## 从字符串生成

安装 `Ofdrw.Net.Layout` 和 `Ofdrw.Net.Packaging`：

```csharp
using Ofdrw.Net.Layout;
using Ofdrw.Net.Packaging;

var options = new PlainTextOfdOptions();
options.Layout.DefaultFontSizeMillimeters = 4;
options.Layout.MaxPageCount = 100;

var converter = new PlainTextOfdConverter(options);
var package = converter.Render("通知 Alpha iii WWW\r\n\r\n尾行");

await using var output = File.Create("notice.ofd");
await new OfdPackageWriter().WriteAsync(package, output);
```

构造函数会验证并快照配置，之后修改 `options` 不影响该转换器。每次调用使用独立的 Flow 状态。

## 从严格 UTF-8 流生成

```csharp
var options = new PlainTextOfdOptions { MaxInputBytes = 100_000 };
var converter = new PlainTextOfdConverter(options);

await using var input = File.OpenRead("notice.txt");
await using var output = File.Create("notice.ofd");
var package = await converter.RenderUtf8Async(input);
await new OfdPackageWriter().WriteAsync(package, output);
```

输入从流的当前位置读取，可以不可 seek；转换器不会关闭输入或输出流。失败或取消后输入位置可能已经前进。

## 文本与分页规则

- UTF-8 解码严格且分块进行，不使用替代字符吞掉错误；只移除输入开头的一个字节 BOM。字符串开头的 U+FEFF 和后续 U+FEFF 都作为内容。
- 拒绝无效 UTF-8、未配对 UTF-16 代理项和 XML 1.0 不允许的标量。
- `CRLF`、`CR`、`LF` 是逻辑换行，`\f` 是换页；中间空行保留，末尾控制符不会额外产生空白尾页。
- 空文本或仅 BOM 的输入生成一张空白页。
- Tab 保留字符并使用 Flow 的四空格推进量，不设置制表位；U+2028/U+2029 作为普通文字。
- 原生文字对象保留准入文字，包括空格、Tab、NBSP 和组合字符；不会添加隐藏原文层。

自动折行和坐标化文本提取意味着读回文本不能保证与原输入逐字节一致，也不能保证恢复原始换行。

## 字体、默认值与预算

默认页面为 A4、四边距 25.4 mm、`SimSun`、黑色、10.5 pt 等值字号。字号和几何配置以毫米表示。

字体族名称只选择 Flow 的既有可移植度量，本入口**不会**测量任意系统字体、嵌入字体、验证字形覆盖或保证跨设备一致外观。需要富文本、表格、图片、页眉页脚、Markdown 或语法高亮时，请使用 [FlowDocument](https://whynpc9.github.io/ofdrw.net/guides/flow-layout/) 等入口。

| 预算 | 默认值 |
| --- | --- |
| `MaxInputBytes`（仅 UTF-8 API） | 4,000,000 字节 |
| `MaxCharacters` | 1,000,000 个 UTF-16 code unit |
| `MaxPageCount` | 10,000 页 |
| `MaxTextElements` | 1,000,000 个文字元素 |

超预算或文字无法放入页面时抛 `InvalidDataException`；取消传播为 `OperationCanceledException`，I/O 错误保持原异常。

## CLI

```sh
ofdrw text-to-ofd notice.txt notice.ofd --font-size-mm 4
ofdrw text-to-ofd --input notice.txt --output notice.ofd \
  --page-width 210 --page-height 297 \
  --margin-left 15 --margin-right 15 --margin-top 15 --margin-bottom 15 \
  --max-input-bytes 100000 --max-characters 50000 \
  --max-pages 100 --max-text-elements 50000
```

使用两个位置路径，或同时使用 `--input` / `--output`，不要混用。`--font-family` 只改变字体名称请求；页面宽高必须成对提供，尺寸和字号必须为正，边距必须非负。输入内容必须是 UTF-8，输出扩展名必须为 `.ofd`。

该命令不支持 stdin/stdout、`--pages`、DOCX 参数或 `--font-directory`。文件在目标目录暂存并原子提交；退出码为 `0`（成功）、`1`（失败）和 `130`（取消）。
