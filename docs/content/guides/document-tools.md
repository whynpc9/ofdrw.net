---
title: "页面重组、水印与附件"
---

CLI 提供页面重排、合并、拆分、水印和签名清理。先安装 [CLI](https://whynpc9.github.io/ofdrw.net/getting-started/installation/)，并用 `ofdrw --help` 核对所安装版本的命令。

```sh
ofdrw reorder input.ofd reordered.ofd --pages 3,1,2
ofdrw split input.ofd selected.ofd --pages 2,1
```

页号从一开始，列表顺序就是输出顺序。API 的页索引从零开始。写入新路径后重新读取输出，确认页序、模板、资源和原文。

## 编辑范围

合并和混合需要证明资源引用闭包，不能安全重映射的未知引用会明确失败。裁剪、签名清理和水印修改可能使原有签名失效；不要把修改后的文件描述为保留原签名效力。

附件逻辑移除与物理载荷清理是不同操作。共享载荷只有在证明没有其他引用时才能删除；详细设计见[附件清理契约](https://github.com/whynpc9/ofdrw.net/blob/main/docs/architecture/r18-attachment-cleanup-design.md)。

多文档容器应先[选择 DocBody](https://whynpc9.github.io/ofdrw.net/guides/multi-docbody/)。类型化对象读写与未知内容保留见[对象编辑](https://whynpc9.github.io/ofdrw.net/guides/standard-object-editing/)。
