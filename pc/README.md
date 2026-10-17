# pc（Scrapling 网页抓取技能）

这是 [Scrapling](https://github.com/D4Vinci/Scrapling) 作者发布的官方 Agent Skill（原名 `scrapling-official`，版本 0.4.15）的副本，只把技能名改成了 `pc`，其余内容与原版一致。

它让 AI 用 Scrapling 抓取网页：普通请求、浏览器渲染、隐身浏览器（可应对 Cloudflare 等反爬）、Spider 批量爬取，以及网页改版后能自动重新定位元素的自适应解析。

## 安装

- **Claude Code**：把整个 `pc` 文件夹复制到 `~/.claude/skills/`，之后用 `/pc` 调用，或者直接让 Claude 抓取网页。
- **Codex**：把 `pc` 文件夹复制到 `~/.codex/skills/`。
- 第一次使用前，在 Python 3.10+ 环境里安装依赖：

  ```
  pip install "scrapling[all]>=0.4.15"
  scrapling install --force
  ```

## 使用须知

只抓取你有权访问的内容，遵守网站的 robots.txt 和服务条款，不要抓取个人敏感信息（详见 [SKILL.md](SKILL.md) 末尾的 Guardrails）。

## 来源与许可证

- 原版：https://github.com/D4Vinci/Scrapling/tree/main/agent-skill/Scrapling-Skill （最新版本以原仓库为准）
- 许可证：BSD 3-Clause，Copyright (c) 2024, Karim shoair，全文见 [LICENSE.txt](LICENSE.txt)。
- 本副本由仓库所有者转载，与 Scrapling 作者无关联，也不代表其认可。
