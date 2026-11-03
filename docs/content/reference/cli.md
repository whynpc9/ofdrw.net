---
title: "CLI 命令与页号"
---

安装 .NET 10 和 `Ofdrw.Net.Cli` 后，使用 `ofdrw --help` 查看当前版本帮助。

```sh
ofdrw convert input.docx output.ofd
ofdrw pdf-to-ofd -i input.pdf -o output.ofd
ofdrw ofd-to-pdf -i input.ofd -o output.pdf
ofdrw ofd-to-svg --input input.ofd --output page.svg --pages 1
ofdrw reorder input.ofd reordered.ofd --pages 3,1,2
```

CLI 页号从一开始，API 索引从零开始。转换后的分页由实际输入、字体及模式确定，不要用 Word 原页码盲选 Native 输出。

命令涵盖文字提取、文本生成、图片/HTML 输出及文档工具；参数以所安装版本帮助为准。当前 main 的 `--native-embedded-bold-behavior throw` 尚未进入 preview.10 CLI，严格字体配置在 preview.10 使用 API。

失败时检查退出码和诊断，原子输出及临时文件边界见[转换约定](https://whynpc9.github.io/ofdrw.net/reference/conversion-contracts/)。
