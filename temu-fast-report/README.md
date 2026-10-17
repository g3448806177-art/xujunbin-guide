# TEMU 快速报单（temu-fast-report）

一个给 Codex 用的 TEMU 半托管报单技能（Agent Skill），目前针对鞋类。
*A Codex agent skill that turns TEMU replenishment sheets into purchase-order workbooks (docs in Chinese).*

把 TEMU「备货管理」导出的 xlsx 交给它，它会：

1. 提取备货表里的 SKU 信息和 SKU 图片；
2. 按**完整 SKC** 查询你自己的「报单记忆」表，命中就直接复用历史采购信息（店铺、网页货号、单价、链接），不再搜图；
3. 没命中的款，到搜款网（sooxie.com）以图搜同款：每款最多两轮、每轮只看前 5 个候选，核对鞋型、颜色、尺码和价格，白名单店铺优先；
4. 生成固定版式的采购单 Excel：左侧「采购」、右侧「网购」，内嵌 SKU 图，另附「人工审核」表；
5. 可选：拍照核对 JIT 标签和报单是否一致。

**报单名**（采购单 F 列和文件名前缀）由每个使用者在安装时自己设置。

## 适用环境

- Windows 11
- Codex（技能安装到 `%USERPROFILE%\.codex\skills`），需要能使用浏览器工具：网页找款要在搜款网上传图片
- Python 3，并安装 `openpyxl`、`Pillow`、`lxml`：

  ```
  python -m pip install openpyxl Pillow lxml
  ```

- 搜款网账号（以图搜款需要登录）

## 安装

1. 打开仓库主页 https://github.com/g3448806177-art/xujunbin-guide ，点 **Code → Download ZIP**，下载后**完整解压**。
2. 进入解压出来的 `temu-fast-report` 文件夹（`install.cmd` 所在的文件夹），双击 `install.cmd`（不需要管理员权限；如弹出安全提示，选择「运行」）。
3. 按提示输入你的**报单名**，例如 `张三`。
4. 重启 Codex，输入 `$temu-fast-report` 使用。

安装程序会：

- 把技能复制到 `%USERPROFILE%\.codex\skills\temu-fast-report`，报单名和桌面路径保存在其中的 `config\settings.json`；
- 在桌面创建空白的 `报单记忆.xlsx` 和 `店铺属性库\店铺属性库.xlsx`（已存在则不覆盖）；
- 重复安装时，先把旧的技能目录备份为 `temu-fast-report.backup-时间戳`。

## 报单名

报单名会写进采购单 F 列，也是文件名前缀。例如报单名为 `张三`，10 月 17 日的第一份采购单就叫 `张三10月17第一次采购单.xlsx`，当天再报单依次是「第二次」「第三次」……

- **修改**：重新运行 `install.cmd`（会显示当前报单名，直接回车就沿用），或用记事本修改 `config\settings.json` 里的 `reporter`。
- **临时换名字**：在对话里告诉 Codex「这次报单名用李四」（技能会以 `--reporter 李四` 运行）。
- 规则：最多 20 个字，不能包含 `\ / : * ? " < > |`。

## 使用

在 Codex 里输入 `$temu-fast-report`，附上备货表，说「报单」即可。还可以：

- 同时提供半托管「订单包裹」导出表，一起找款（SKC 列标记为「（半托管订单）」）；
- 发 JIT 标签照片，让它和报单逐条核对。

交付时会汇报匹配、网购、待审核、搜索失败的数量，以及「记忆待人工录入」。**下单前请看一遍「人工审核」表**：自动匹配不等于人工确认。

### 报单记忆（只读）

- 位置：桌面 `报单记忆.xlsx`。新安装是空表，不会命中任何 SKC。
- 技能**只读**这张表，不会自动写入。每批报单核对无误后，由你手动把要保留的记录录入（按空白表的表头布局填写）。
- 之后再遇到同一个完整 SKC，就直接复用历史记录，不再搜图。

### 店铺属性库

- 位置：桌面 `店铺属性库` 文件夹。把 TEMU 导出的商品属性表（含 `SKC_ID`、`SKC货号`、`SKU货号`、`商品属性` 列）放进去，或者手动填写 `店铺属性库.xlsx`。
- 采购单的「商品属性」列按 SKC 精确匹配，输出鞋面、副鞋面、内里、鞋底、鞋垫五行材料；查不到就留空，不猜。

### 白名单店铺

`temu-fast-report/references/preferred-shops.json` 是搜款网白名单：候选里出现白名单店铺并且核对通过时优先选它。仓库自带的是作者在用的鞋类店铺，可以换成你自己的供应商。`SKILL.md` 里的店铺别名规则（如「潮旭网批=旭日升一号」）也可以按需修改。

## 采购单格式

| 区域 | 列 | 内容 |
|---|---|---|
| 采购（左） | A–L | 码数、颜色、备货信息货号、成本、件数、报单名、店铺、地址、图片、网站备注、商品属性、SKC |
| 分隔 | M | 空白 |
| 网购（右） | N–S | 码数、颜色、备货信息货号、图片、商品属性、SKC |

第 2 个工作表「人工审核」只列网页搜款成功的记录：目标 SKU 图、确定的店铺商品主图、鞋店名称。完整规则见 [format-spec.md](temu-fast-report/references/format-spec.md)。

## 目录结构

```
├── install.cmd / install.ps1   Windows 安装程序
├── 安装说明.txt
└── temu-fast-report/           技能本体
    ├── SKILL.md                技能说明（Codex 读取）
    ├── agents/openai.yaml      Codex 界面显示信息
    ├── assets/                 采购单模板、空白记忆表、空白属性库
    ├── references/             格式规范、记忆操作说明、白名单
    └── scripts/                prepare / build / 记忆查询 / 候选拼图 / 标签核对
```

## 手动安装（不用安装程序）

把 `temu-fast-report` 文件夹复制到 Codex 的技能目录，然后新建 `temu-fast-report/config/settings.json`：

```json
{
  "reporter": "张三",
  "memory": "C:\\Users\\你的用户名\\Desktop\\报单记忆.xlsx",
  "attribute_dir": "C:\\Users\\你的用户名\\Desktop\\店铺属性库"
}
```

`memory` 和 `attribute_dir` 不写时，默认用桌面上的 `报单记忆.xlsx` 和 `店铺属性库`。

## 注意事项

- 本项目与 TEMU、搜款网没有任何关联。自动找款请遵守网站规则；技能内置了提交间隔和限流冷却。
- 不要把你自己的 `报单记忆.xlsx`、属性库或报单结果提交到公开仓库（`.gitignore` 已默认排除）。

## 许可证

[MIT](LICENSE)
