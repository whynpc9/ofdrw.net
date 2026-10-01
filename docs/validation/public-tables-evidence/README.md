# 可下载的实际表格验收证据

这里保存了本会话实际生成、检查的原始文件；从干净 checkout 即可查看。没有重新构造 OFD/PDF，也没有修改页面、字体或 ZIP 时间戳来满足记录的哈希。

最新实际 Preview 源码为 `31f0b75fc6b2846b6651591cf006834b91ae7f4e`；最新功能验证源码为 `26f33e91d716c863bed09e5feb3c95a0a423b00f`。两者之间仅非法跨度的失败消息改变，无受影响的正常页面；完整范围见[验收记录](../public-tables-2026-10-02.md)。

直接查看最新的五份 OFD 导出 PDF（macOS Preview 实际逐页检查11页后关闭窗口）：

| PDF | 实际查看页 | 页面图 |
| --- | --- | --- |
| [公开 Table](preview/tables-public.pdf) | 1–3 | [1](preview/pages/tables-public-1.png)、[2](preview/pages/tables-public-2.png)、[3](preview/pages/tables-public-3.png) |
| [DOCX 显式 Native 表格](preview/tables-native.pdf) | 1–2 | [1](preview/pages/tables-native-1.png)、[2](preview/pages/tables-native-2.png) |
| [DOCX 默认表格](preview/tables-default.pdf) | 1–2 | [1](preview/pages/tables-default-1.png)、[2](preview/pages/tables-default-2.png) |
| [既有 DOCX 显式 Native 基准](preview/baseline-native.pdf) | 1–2 | [1](preview/pages/baseline-native-1.png)、[2](preview/pages/baseline-native-2.png) |
| [既有 DOCX 默认基准](preview/baseline-default.pdf) | 1–2 | [1](preview/pages/baseline-default-1.png)、[2](preview/pages/baseline-default-2.png) |

[实际页码/环境/源码记录](preview/acceptance.json)、[本轮文件哈希清单](preview/artifact-manifest.tsv)也可直接读取。PNG是辅助证据，不能代替上面的实际Preview过程。

[原始完整产品包](products.tar.zst)保存 `initial/`（含斜线缺陷）、`fixed/`、`review1/current/` 三次生成的全部原始DOCX/OFD/PDF/PNG、记录和哈希；另含独立验证日志、七份最终300测试TRX、11包消费清单与负向像素复现源码。档案内路径 `public-tables/...` 对应验收文中的本地 `artifacts/public-tables/...`。历史快照也保留，因此记录中的旧哈希可以逐个复核。

档案约16 MB，展开约189 MB；用zstd的长窗口压缩重复的OFD内嵌字体载荷，展开后每个原始文件逐字节还原。[档案清单](bundle-manifest.json)列出108个文件的大小与SHA-256，并记录压缩包自身的SHA-256。

需要Python3和zstd；macOS可用 `brew install zstd`，Ubuntu可用 `sudo apt-get install zstd`。在仓库根目录运行：

```sh
python3 scripts/verify-public-table-evidence.py
python3 scripts/verify-public-table-evidence.py --extract /tmp/ofdrw-table-evidence-review
```

第二条命令同时验证和提取，拒绝覆盖已有文件。随后可直接打开 `/tmp/ofdrw-table-evidence-review/public-tables/review1/current/*.ofd`，核对其哈希及OFD原文；最新PDF/PNG的独立副本也由同一脚本核对。CI会验证该档案及全部直接查看副本的完整性。
