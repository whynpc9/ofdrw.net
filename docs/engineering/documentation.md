# 文档维护与发布

网站正文唯一来源为 docs/content/，Nimbus 工程在 docs/site/。docs/ 根目录只放分类目录。

```sh
cd docs/site
npm ci
npm run check
npm run build
npm run preview -- --host 127.0.0.1
```

Node 要求 >=22.12；使用 Node 24。依赖锁定并提交 lockfile，升级 Nimbus 时遵循 site/AGENT.md。正文使用 Markdown 和 title frontmatter；需要组件时才用 MDX。

PR 构建检查，main 更新后通过 docs-pages.yml 发布 GitHub Pages。站点地址 https://whynpc9.github.io/ofdrw.net/，base 为 /ofdrw.net/。不启用需要服务端的 MCP 或协商响应。文档 CI 不下载历史 LFS；必要字体由 npm 锁定的依赖提供，不重复提交字体载荷。

新功能必须标明首次可用版本；尚未发布的 main 功能单独标记。历史验收 JSON 和 manifest 不随迁移重新签收；Q05 仍拒绝与当前候选不匹配的旧指纹。证据路径按原提交解释。

每次迁移检查相对链接、包内文档、脚本路径和 storage gate；浏览器检查桌面/窄屏导航、中文搜索、代码复制及深层链接。
