# SKC报单记忆操作

需要手动查询、图片提取、人工录入或匹配字段时才读本页；平常执行prepare/build即可。

## 查询与图片

```bash
python3 scripts/memory_store.py lookup --records work/records.json --out work/memory-hits.json
python3 scripts/memory_store.py image --skc 26341603085 --ref 'Sheet1!I5' --out work/memory-images/skc-26341603085
```

默认记忆为`Win11 桌面/报单记忆.xlsx（实际路径见技能config/settings.json）`。`--memory`用于用户指定路径或隔离测试；查询的`--cache`指定SQLite缓存。默认缓存为原表旁`.报单记忆.skc.sqlite`，仅有版本、文件签名和`positions(skc,sheet,row,section)`，不存采购详情或图片。采购／网购SKC列由第2行的两个`SKC`表头识别，兼容手工维护的`H/L`布局和原`I/M`布局，不以固定列号判断。

根据路径、大小、纳秒修改/变更时间和文件标识检查缓存，变化/缺失/损坏就重建。原表不可用时警告并转网页找款，不使用旧记录。批量SKC去重后参数化SQL精确查询；未命中不取详情。并发修改时重试一次，仍变化则停止使用记忆。

SKC字符串仅去首尾空格，保留前导零；Excel整数转数字字符串。小数、布尔、科学计数法文本、带字母或货号的内容不作为SKC。超出Excel可靠精度的数字不猜补，长SKC应以文本保存。

`hits`按命中SKC分组，每条有`ref`、`kind`、`status`、`imageRef`及必要采购字段。`memoryCopyByRecord`只返回每条本次记录的优先复用项；多条历史时优先有采购链接、核验时间最新的一条，完整候选仍保留在`hits`。完整SKC命中后不读链接、不读历史图、不上传搜图：采购项照抄店铺、网页货号、单价和链接，已有地址一并复用；地址或链接为空就留空，不打开历史链接或网页补查；只替换本次SKU货号、颜色、码数、数量；网站备注写`历史网页货号*本次颜色*本次码数*本次数量`。网购项直接放网购区。只有`webRecordIds`中的未命中SKC网页搜索。

命令默认只打印摘要和输出路径。需要某一命中时，本地读取`hits[skc]`，不传入整本、整索引或图片编码。`summary.resultBytes`是UTF-8字节数，不是精确词元数。

图片的`ref`必须来自本次命中，不能猜行号。只有用户明确要求查看历史图片时才提取；正式报单不使用历史图，继续使用本次SKU图。提取前重新校验当前位置仍为该SKC，再读取这一条WPS图片。返回路径使用原图片扩展名。不支持的图片存储或资源缺失会报错；保留本次SKU图，不读取其他图片补查。

## 匹配字段示例

```json
{
  "date": "0913",
  "matched": {
    "文件标识:2": {
      "shop": "示例店铺", "webNo": "A1", "webColor": "黑色", "cost": 25,
      "url": "https://example.com/item/1", "whitelist": false,
      "verification": {
        "style": true, "color": true, "size": true, "price": true,
        "checkedAt": "2026-09-13T20:00:00+08:00", "reason": "实际观察到的核对依据"
      },
      "candidates": [{"shop": "示例店铺", "webNo": "A1", "image": "/abs/candidate.jpg", "chosen": true}]
    }
  },
  "unmatched": ["文件标识:3"],
  "searchEvidence": {
    "文件标识:3": {"rounds": [
      {"kind": "sku", "loaded": true, "candidateCount": 5, "outcome": "no_match", "reason": "前5排除原因"},
      {"kind": "main", "loaded": true, "candidateCount": 5, "outcome": "no_match", "reason": "前5排除原因"}
    ]}
  },
  "pending": {"文件标识:4": {"reason": "关键部位无法确认"}},
  "search_failed": {"文件标识:5": {"errorCode": "upload_not_attached", "reason": "文件选择完成后网页仍显示点击上传，文件未挂载到上传控件"}},
  "unmatchedCandidates": {}
}
```

`reporter`可省略：build使用prepare时确定的报单名；如填写，必须与prepare的报单名一致，否则build停止。示例ID、商品字段、时间和依据须替换为本批实际值。核验值来自当前观察，不能为通过检查补true。`cost`为网页单价，数量用原表；候选图片必须存在。新搜款不获取地址，`address`可省略或留空；历史地址有值则直接复用。时效字段不再获取或写入批注。无主图时只有sku轮，有主图时两轮完成才可归网购；疑似但不能确认用pending。旧源行号仅在唯一对应原记录时兼容。新的`search_failed`必须同时写`errorCode`和实际`reason`；允许值为`upload_unavailable`、`upload_not_attached`、`upload_connection_lost`、`search_failed`、`rate_limited`、`login_failed`。系统文件窗口看似选完但网页仍显示“点击上传”时用`upload_not_attached`，不能写`unmatched`。

## 人工录入记忆

记忆库是只读查询源。build回读验证正式报单后，仅在报单旁保存`.memory-sync.json`，其中`status=manual_required`，说明本批**没有**新增、更新、删除或同步任何记忆记录。`memory_store.py sync`已禁用；不能用脚本补写、重试续写或批量回填。

人工完成商品、颜色、尺码、当前价格和供应商核对后，直接在`Win11 桌面/报单记忆.xlsx（实际路径见技能config/settings.json）`中添加或修改需要保留的记录，并保持原表布局和WPS单元格图。是否更新同SKC同供应商同规格、是否追加不同供应商、以及是否保留待审核或搜索失败记录，均由人工决定。缺SKC的行不参与后续查询索引。

原表是唯一数据源。人工修改或删除后，下一批查询会因文件签名变化重建索引；不会从旧报单、旧缓存或备份补回记录。`记忆审核`不进入可复用索引。
