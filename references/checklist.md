# 症状 → 原因 → 处理 速查表

先按症状定位，再去看 `pitfalls.md` 里对应的编号。

## 双击 / 编辑

| 症状 | 最可能原因 | 处理 |
|---|---|---|
| 双击 OLE 没反应或空白 | 用独立 OPJU 替换了 `Contents` 流 | 用 `copy_page('OLE')` → WPS `PasteSpecial(10)` 重新生成工厂 |
| PowerPoint 报 `OLEFormat: Object does not exist` | 手工重建 OLE 复合文件时写了时间戳 | 不要手工重建；改用工厂 OLE |
| 双击后 Origin 打开的是空工程 | `p:embed` 丢失或 rId 指向错误 | 检查 slide rels 与 `p:oleObj r:id` |
| `op.save()` 报 `must supplied a path` | 在嵌入工程里保存 | 不要保存嵌入工程；另存独立 OPJU 后再生成工厂 |
| `op.exit()` 后 WPS/PowerPoint 的 ActivePresentation 失效 | Origin 退出断开了 OLE 连接 | 工厂生成脚本里最后再退出 Origin，且不要依赖已关闭的演示对象 |

## 尺寸 / 位置

| 症状 | 最可能原因 | 处理 |
|---|---|---|
| 图只占框的 ~83%，四周白边 | 预览含页边距 + 长宽比不匹配 | 先看 `ole_inspect.py frames/factory`；必要时用 ExportClip 生成（注意 #17） |
| 图被压扁或拉长 | 预览长宽比与图框差异大 | 检查工厂里是否混入页外对象；或为不同图框分别生成工厂 |
| 标题/图例被裁切 | 预览边界比图框小 | 同上；优先保证内容完整，再调图框 |
| 标题压住刻度 | K 刻度变长，标题位置没变 | 下移 `GraphObjects['XB'].Top`，可配合缩小 `x.label.pt`；逐页验证 |
| 标题跑到图框外、压到相邻图 | 标题下移过多，且图框太矮 | 该页单独用“不下移”的工厂，或调整版式 |
| 一张图在两个页面大小不一致 | 两页图框尺寸不同，或 imgW/imgH 不同 | 以图框为准分别验收；必要时分别生成工厂 |
| 图位置整体偏移 | 替换时改了 `a:off/a:ext` | 只替换 bin/media，不动形状几何；编辑 `p:oleObj` 时只改 imgW/imgH |

## 显示 / 重影

| 症状 | 最可能原因 | 处理 |
|---|---|---|
| 图上叠了一层半透明图片 | `p:oleObj` 内置的 `<p:pic>` 预览被 WPS 同时显示 | `strip_ole_previews()` 删除 OLE 内部的 `p:pic` |
| WPS 整页斜线 | 走了 VML `mc:Choice`，WMF 预览损坏 | 删除该 rId 周围的 `<mc:Choice>...</mc:Choice>` |
| 页面上的真实图片被误删 | 正则删到了页面的 `p:pic` | 只删 `p:oleObj` 元素内部的 `p:pic` |
| 图例多出 `EnvLow`/拟合条目 | 自动图例被拟合 plot 更新 | 删除自动 Legend，改用静态 label |
| 工厂预览里多出一个孤立图例 | 页外残留对象被 ExportClip 导出 | 在 Origin 里 `GraphObjects['Legend'].Destroy()` 后重新生成 |

## 数据 / 轴

| 症状 | 最可能原因 | 处理 |
|---|---|---|
| 读不到图里的数据 | 用了 `DataRange.GetData()` | 改用 LabTalk `range -wx/-wy` |
| `op.open()` 返回 False | 旧工程常见，实际已加载 | 以 `op.graph_list()` 为准 |
| 轴标签没变 | 改了 `layer.x.title$` | 用 `ax.title` 和 `layer.x.label.formula$` |
| 手绘线与曲线错位 | 改了轴范围/数据 | 只改显示公式，不动范围 |

## 流程 / 效率

| 症状 | 最可能原因 | 处理 |
|---|---|---|
| 反复渲染大 PPT，上下文爆掉 | 每次都整页导出 + 读图 | 先用 `ole_inspect.py` 在包内判断；最后只做一轮整页验收 |
| 替换后多个页面串图 | 多个 OLE 共享同一个 `imageNN.emf` | 先解析 media 引用关系，共享媒体要统一工厂 |
| 替换后其他页变了 | 误改了共享的 media 或全局 XML | 每次替换前列出 embedding ↔ media 映射，只动目标项 |
