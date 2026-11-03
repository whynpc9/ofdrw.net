# 文档清理与 Nimbus 站点验收（2026-11-03）

基线：main `2489966c47833cb137fe33cdc4c3d2df60daec3a`。逐文件处置见 [迁移清单](../engineering/documentation-migration.md)。本轮删除全部 24 篇旧教程，根目录正文分类迁移，网站正文唯一来源为 `docs/content/`，工程在 `docs/site/`。

- Node 24.19.0；Nimbus 0.19.2、Astro 7.3.3，实际 npm 依赖固定并提交 lockfile。全新 npm ci、静态构建及 Astro 类型检查通过（0 errors / 0 warnings）。
- 43 篇正文，生成 44 个 HTML 页面。自有检查器核对实际构建结果的内部路径、base 和锚点；全部通过。Pagefind 索引 43 页。
- Nimbus check 的环境、结构、迁移和 TypeScript 项通过；它不 lint `.md` 正文，因此 Authoring 明确为 partial。未把该结果表述为完整内容审校。
- 44 项 Python 工程脚本测试通过，1 项既有条件跳过；actionlint 通过。全解决方案 Release pack 通过；默认和可选包内要求的文档条目均存在，许可证和依赖清单检查通过。
- 新指南中完整 C# 示例以本次本地包编译通过，0 警告 / 0 错误。未运行转换或重新生成视觉产品，本轮不产生新的 DOCX/OFD/PDF 页面验收结论。
- Q05、API 文档基线和第三方基线只迁移路径，内容与原提交逐字节一致。证据 JSON、manifest、哈希和原视觉证据不重写、不重新签收；旧 Q05 不放行本轮或新发布。
- 暂存后存储检查通过，包括 e2e/tests 载荷检查。没有新增源码字体载荷，站点字体由锁定 npm 依赖提供；站点构建不下载历史 LFS。

## 浏览器验收

按 AGENTS.md 使用 `gpt-6.1-terra` High 子代理独占 ego-browser。实际检查约 1240×949 桌面与 390×844 窄屏：首页、导航、中文搜索“字体”和“表格”及结果跳转、深层页重载、主题切换、复制按钮状态、宽代码和表格的横向滚动。发现重复 base 和移动菜单默认云朵后已修复并补验。主代理查看移动菜单及字体代码截图，确认修复。

复制只验证按钮显示 Copied!，未检查剪贴板字节；横滚使用 DOM 设置，未检查真实触屏手势。范围仅为上述样本及视口。子代理关闭其浏览器任务空间。

详细日志、包、示例编译及 UI 截图位于 ignored `artifacts/docs-migration/`；必要结论保存在本报告。线上部署由独立 `docs-pages.yml` 执行，检查其当前运行结果；本报告的本地验收不证明远端部署成功。
