# Origin PPT OLE Skill

一套面向 **Windows + Origin/OriginPro + WPS + python-originpro** 的 Skill / 工具箱，用于批量处理 PowerPoint 里嵌入的 Origin 图。

它解决的不是“画图”，而是**把改好的 Origin 图安全地放回 PPT**，并且：

- 双击仍然能打开 Origin 编辑；
- 图的大小、位置、长宽比与原图框一致；
- 不会在图上叠出一层预览图；
- 不会因为一个图例或一次 `ExportClip` 把整页搞坏。

> 这些结论来自真实的 100 MB+ 论文汇报 PPT 的反复翻车与修复。先读 [`references/pitfalls.md`](references/pitfalls.md)，能省掉大部分试错。

---

## 解决什么问题

| 场景 | 能力 |
|---|---|
| 把 Origin 图的 ℃ 轴改成 K | 只改刻度显示公式，不动数据和轴范围，避免 PPT 手绘标注错位 |
| 从图里读回原始数据 | LabTalk `range -wx/-wy` 的可靠读法 |
| logistic / 指数 / 高斯 / 多项式拟合 | 拟合曲线作为原生 plot 加回图里 |
| 多组数据填充包络 | 颜色索引、`set %C -pfv/-cf` 的正确写法 |
| 替换 PPT 里的 OLE | 工厂 OLE + 正确的 rId/media 映射 |
| 图变小 / 有白边 | 用 `imgW/imgH` 与图框长宽比定位，不靠猜 |
| 图上叠了一层图片 | 删除 `p:oleObj` 内部的 `<p:pic>` 预览 |
| 双击空白 / 打不开 | 避免“独立 OPJU 直接写回 Contents”的致命错误 |
| 大文件反复渲染太慢 | 包内诊断 + EMF 预览，最后只做一轮整页验收 |

---

## 安装

作为 Claude Code Skill：

```bash
git clone https://github.com/ArthurEccleston/origin-ppt-ole-skill.git \
  ~/.claude/skills/origin-ppt-ole-skill
```

然后在 Claude Code 里说“用 origin-ppt-ole-skill 处理这个 PPT 的 Origin 图”，或直接提“Origin / OLE / 双击打不开 / 图变小”。

作为普通 Python 工具：

```bash
pip install originpro pywin32 olefile Pillow numpy scipy
```

> Origin 需要本机安装并授权；`originpro` 只是 COM 封装。
> WPS 通过 `KWPP.Application` 驱动，不需要写死安装路径。

---

## 快速开始

```bash
# 1. 看目标 PPT 里所有 OLE 的映射与尺寸
python scripts/ole_inspect.py frames deck.pptx

# 2. 看工厂 OLE 的原生尺寸 / EMF 物理尺寸 / 长宽比
python scripts/ole_inspect.py factory factory.pptx

# 3. 直接把工厂预览转成 PNG（不打开 PPT）
python scripts/ole_inspect.py preview factory.pptx factory.png

# 4. 去掉 Origin 图上的预览重影
python scripts/ole_inspect.py strip deck.pptx deck_no_preview.pptx
```

Python 里：

```python
import ppt_ole_ops as po

# 生成合法工厂 OLE：Origin copy_page('OLE') -> WPS PasteSpecial(10)
emb, media = po.make_factory_ole(r"graph_K.opju", r"factory.pptx")

# 映射并替换（slide/rId/embedding/fallback 全部自动解析）
for o in po.list_ole_objects(r"deck.pptx"):
    print(o['slide'], o['rid'], o['embedding'], o['fallback_media'])

po.replace_ole(
    "deck.pptx",
    "ppt/embeddings/oleObject1.bin",
    "factory.pptx",
    "ppt/media/image11.emf",
    "deck_out.pptx",
    remove_vml=True,
)

# 验证：WPS 双击 -> Origin，检查图、plot 数、X 轴标题
print(po.verify_ole("deck_out.pptx", 3))
```

---

## 三条必须先知道的事实

1. **独立 `.opju` ≠ 嵌入 Origin 文档。**
   把独立 OPJU 写回 OLE 的 `Contents`，PPT 上看起来正常，双击却打不开。必须用 `copy_page('OLE')` → `PasteSpecial(10)` 生成工厂 OLE。

2. **`oleObjectN` 编号与幻灯片顺序无关。**
   解析 `ppt/slides/_rels/slideN.xml.rels`，按 `rId → oleObjectN.bin` 映射；同一张图还可能被多个页面共享同一个 EMF。

3. **`SetExportClip(True)` 不是“裁掉页边距”。**
   它会把页外对象也导出。页外残留的 Legend 会让预览长宽比从 1.3:1 变成 2.0:1，替换后图被压扁。先清页外对象，再生成工厂。

---

## 尺寸 / 重影排查顺序

```text
图在框里只占 83%
  ├─ ole_inspect.py frames   -> 看 imgW/imgH 与 frame 的长宽比差
  ├─ < 3%  -> 主要是页边距；用 ExportClip 工厂（确认没有页外对象）
  └─ 很大  -> 查工厂是否混入页外图例（GraphObjects 里 Left 远大于页宽）

图上叠了一层图片
  └─ ole_inspect.py strip    -> 删除 p:oleObj 内部的 p:pic

标题压住 K 刻度
  ├─ GraphObjects['XB'].Top += 60
  ├─ x.label.pt 16 -> 14
  └─ 逐页渲染：小图框页可能放不下下移后的标题，需要单独工厂
```

更完整的“症状 → 原因 → 处理”见 [`references/checklist.md`](references/checklist.md)。

---

## 目录

```text
.
├── SKILL.md                     # Skill 定义与工作流
├── scripts/
│   ├── origin_ops.py            # OPJU：读数据、改轴、拟合、包络、图例
│   ├── ppt_ole_ops.py           # PPT/WPS：OLE 映射、工厂、替换、验证
│   └── ole_inspect.py           # 包内诊断、EMF 预览、去除预览图
├── references/
│   ├── pitfalls.md              # 21 条踩坑记录
│   ├── checklist.md             # 症状速查表
│   └── diagnostic-prompt.md     # 可直接粘贴的排障提示词
└── README.md
```

---

## 环境与限制

- 仅在 **Windows 10/11 + Origin 2024/2025b + WPS Office + Python 3.11+** 上验证。
- OLE 的显示与激活依赖本机注册的 Origin OLE 服务器；换机器后必须重新双击验证。
- WPS 与 PowerPoint 对 `mc:Choice` / `mc:Fallback` 的选择不同；本工具默认删除 VML 分支、强制走 EMF/嵌入预览。
- 所有脚本默认在副本上操作，不修改原文件。

## License

MIT
