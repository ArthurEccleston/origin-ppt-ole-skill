---
name: origin-ppt-ole-skill
description: "Origin/OPJU 与 PPT 内嵌 OLE 图的批处理工具。读取 Origin 图里的原始数据、Celsius→Kelvin 坐标轴转换、logistic/指数/高斯/多项式拟合、生成拟合包络、把图作为合法 Origin OLE 替换回 PPT，并修复“双击打不开 / 图变小 / 图上叠了一层图片 / 轴标题压刻度”等问题。当用户提到 Origin、OriginPro、OPJU、OLE、PPT 里的 Origin 图、坐标轴单位、拟合曲线、包络、双击不显示图、OLE 重影时使用。"
---

# Origin / PPT OLE 操作

面向 **Windows + Origin + WPS + python-originpro** 的一套可复用流程。核心目标只有一句话：

> **替换进 PPT 的图，必须仍然是“真正的嵌入 Origin 对象”——能双击编辑，同时视觉尺寸和位置与原图一致。**

## 环境

- Windows 10/11
- Origin / OriginPro 2024 或 2025b
- Python 3.11+，安装：
  - `originpro`（官方 Python 包）
  - `pywin32`
  - `olefile`
  - `Pillow`
  - `numpy` / `scipy`（拟合用）
- WPS Office（用于打开/验证 PPT；脚本通过 `KWPP.Application` 驱动，不依赖安装路径）
- 若本机同时装了 PowerPoint，验证仍然以 WPS 为准（两者对 OLE 的 VML/EMF 分支选择不同）

## 四条铁律

1. **不要把独立 `.opju` 直接塞回 OLE 的 `Contents` 流。**
   - PPT 里的 Origin 图是“嵌入 Origin 文档”，不是普通 OPJU。
   - 独立 OPJU 替换后，幻灯片回退图看起来正常，但双击时 Origin 不认，空白。
   - 正确做法：Origin 里 `gp.copy_page('OLE')` → WPS `Shapes.PasteSpecial(10)` 生成真正的嵌入 OLE，再替换。

2. **不要假设 `oleObjectN` 编号和幻灯片顺序一致。**
   - 实际文件可能有 19+ 个嵌入对象，编号不连续、跨页复用。
   - 必须解析 `ppt/slides/slideN.xml.rels`，按 `rId → oleObjectN.bin` 映射。

3. **验证看渲染结果，不看“双击能开”一项。**
   - OLE 能打开只说明嵌入有效；图变小、标题被压、叠了预览图，都需要逐页导出 PNG 目视。
   - 但也不要每次都渲染整份大文件——先用 `ole_inspect.py` 在包内把问题筛掉。

4. **永远在副本上操作。**
   - 原 PPT 保持不动；所有替换、删除、移动都在新文件上完成。

## 目录

```text
scripts/
  origin_ops.py     # OPJU 侧：读数据、改轴、拟合、包络、图例
  ppt_ole_ops.py    # PPT/WPS 侧：OLE 映射、工厂 OLE、替换、验证
  ole_inspect.py    # 包内诊断：帧/原生尺寸长宽比、EMF 预览、去除预览图
references/
  pitfalls.md       # 全部踩坑记录（先读这个）
  checklist.md      # 症状 → 原因 → 处理 速查表
  diagnostic-prompt.md  # 丢给任意 AI 的排障提示词
```

## 常用工作流

### A. 摄氏度轴改成开尔文

```python
import origin_ops as oo
oo.connect_standalone()
gp = oo.open_project(r"...\graph.opju")
gl = gp[0]; gl.activate()
oo.set_axis_kelvin(gl, 'x')   # title -> Temperature (K), ticks x+273.15
oo.save_export(gp, r"...\graph_K.opju", r"...\graph_K.png")
oo.exit_origin()
```

原理：数据仍保持 Celsius，只改刻度显示 `layer.x.label.formula$ = "x+273.15"`；不改数据、不改轴范围，PPT 上已有的手绘标注不会错位。

CLI：

```bash
python scripts/origin_ops.py axis --opju graph.opju --axis x --out graph_K.opju
```

### B. 读取图里的原始数据

图里的数据是数据集引用，不是副本。可靠写法：

```labtalk
range -wx rx = 1;
range -wy ry = 1;
```

```python
xy = oo.read_plot_data(gl, 1)     # 1-based plot index
all_data = oo.read_all_plots(gp)
```

### C. 拟合 + 添加拟合曲线

```python
xy = oo.read_plot_data(gl, 1)
x = [p[0] for p in xy]; y = [p[1] for p in xy]
fit = oo.fit_series(x, y, 'logistic')
wks = oo.make_fit_sheet(fit['xfit'], {'GSD': fit['yfit']})
oo.add_fit_line(gl, wks, 1, color=(255, 0, 0), width=2)
```

| 模型 | 典型用途 |
|---|---|
| `logistic` | 硬度/软化曲线（S 形） |
| `exp_growth` | 晶粒尺寸随温度增长 |
| `exp_decay` / `poly2` | 驱动压力趋势 |
| `gaussian` | 峰值型数据（f_domain、Zener pinning） |
| `poly2` / `poly3` | 平滑趋势 |

### D. 生成合法 OLE 并替换进 PPT

```python
import ppt_ole_ops as po

# 1) 工厂：只含一个真正的嵌入 Origin 对象
emb, media = po.make_factory_ole(r"...\graph_fit.opju", r"...\factory.pptx")

# 2) 目标映射（slide / rId / embedding / fallback media）
objs = po.list_ole_objects(r"...\deck.pptx")

# 3) 替换 OLE bin + EMF 回退图，并去掉 VML 分支
po.replace_ole(deck, 'ppt/embeddings/oleObject1.bin', factory,
               'ppt/media/image11.emf', out_deck, remove_vml=True)
```

### E. 图变大/变小/叠影：三个独立问题

这三个症状经常被混为一谈，必须分开处理。

#### E1. 图在框里只占 ~83%（四周有白边）

原因：工厂 OLE 的原生预览带整页边距，且预览长宽比与幻灯片图框不一致；WPS 按长宽比适配后，绘图区缩小。

处理顺序：

1. 生成工厂时尝试：
   ```python
   gp.obj.SetExportClip(True)
   gp.obj.SetExportMargin(0)
   gp.obj.SetExportRatio(100)
   gp.copy_page('OLE', 300, 100)
   ```
2. 用 `ole_inspect.py factory` 看 EMF 长宽比，用 `frames` 看目标图框长宽比。
   - 差值 < ~3%：通常能填满。
   - 差值很大（例如 2.01 vs 1.24）：**不要用**，先查工厂里是否混入了页外对象。
3. `ExportClip` 的真实语义不是“裁掉页边距”，而是“把页外内容也导出”。页外残留的 Legend/Text 会把长宽比拉爆。先清掉它们再生成。

#### E2. 工厂预览里多出页外图例

```python
gp = op.graph_list()[0]; gp.activate()
gos = gp[0].obj.GraphObjects          # 按名称索引，可迭代
for go in gos:
    print(go.Name, go.Left, go.Top, go.Width, go.Height)   # 单位 1/1000 inch
stray = gos['Legend']
stray.Destroy()                       # 没有 .Delete()
```

删除后重新生成工厂。不要改原 OPJU 的可见内容。

#### E3. 图上叠了一层图片（重影）

PPTX 的每个 `p:oleObj` 里通常带一个 `<p:pic>` 预览图。WPS 有时会同时显示 OLE 和这张图，看起来像半透明重影。

```python
from ole_inspect import strip_ole_previews
strip_ole_previews('deck.pptx', 'deck_no_preview.pptx')
```

删除后 OLE 仍然从嵌入的 Origin 文档渲染，双击编辑不受影响。注意：只删 `p:oleObj` **内部**的 `p:pic`，不要碰页面上真正的图片（SEM 照片等）。

### F. 逐页验证

```python
res = po.verify_ole(out_deck, 3)
# {'ok': True, 'graphs': ['Graph1'], 'plots': 6, 'x_title': 'Temperature (K)'}
po.render_slide(out_deck, 3, r"...\slide3.png")
```

验收清单：

- [ ] WPS 双击 OLE → Origin 出现图，plot 数量正确
- [ ] X 轴标题含 `(K)`
- [ ] 渲染图里图填满原框，轴标题/图例未被裁切或压住
- [ ] 页面上没有第二层预览图
- [ ] 手绘线/I/II/EXO/图片与文字位置未变
- [ ] 非目标页未变化
- [ ] 原文件未被修改

## 高效排障（重要）

不要为了看一个小问题反复渲染 100 MB 的 PPT，那样既慢又容易爆上下文。优先在包内判断：

```bash
python scripts/ole_inspect.py frames deck.pptx          # 所有 OLE 的帧/原生尺寸/长宽比差
python scripts/ole_inspect.py factory factory.pptx      # 工厂原生尺寸 + EMF 物理尺寸
python scripts/ole_inspect.py preview factory.pptx p.png  # 直接把 EMF 渲染成 PNG 看内容
```

只有在候选方案确定后，才做一轮整页渲染验收。

## 文件

- `scripts/origin_ops.py` — OPJU：读数据、改轴、拟合、包络、图例
- `scripts/ppt_ole_ops.py` — PPT/WPS：OLE 映射、工厂、替换、验证
- `scripts/ole_inspect.py` — 包内诊断、EMF 预览、去除 OLE 预览图
- `references/pitfalls.md` — 全部坑与复现条件
- `references/checklist.md` — 症状速查
- `references/diagnostic-prompt.md` — 可直接粘贴的排障提示词
