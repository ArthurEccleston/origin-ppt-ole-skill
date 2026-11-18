# Origin / PPT OLE 踩坑记录

这份记录来自多个真实 PPT 的处理（Origin 图嵌入、批量化改轴、替换 OLE），目的是让下一次不再重复同样的错误。

## 1. 最核心的坑：独立 OPJU ≠ 嵌入 Origin 文档

PPT 里双击能打开的 Origin 图，其 OLE `Contents` 流是一个**嵌入 Origin 文档**。把它解出来、用 `originpro` 改完、再 `op.save()` 成普通 `.opju`，看起来是同一种文件（都以 `CPYUA` 开头），但**不是同一种文档类型**。

后果：
- 幻灯片上的回退图/EMF 仍然显示，像素比对也能通过；
- 双击时 Origin 的 OLE 服务器不认这个独立工程，窗口不出现或空白；
- 原稿能双击，改稿不能。

正确做法：
1. 在 Origin 里打开改好的 OPJU；
2. `gp.copy_page('OLE', 300, 100)`；
3. 在 WPS 里新建空白页，`Shapes.PasteSpecial(10)`；
4. 保存这个“工厂 PPTX”，从中取 `ppt/embeddings/oleObject1.bin` 和 `ppt/media/image1.emf`；
5. 用它替换原 PPT 的目标 OLE bin 和 EMF 回退图。

## 2. OLE 编号不是连续的 1..N

实际文件可能有 19+ 个 `oleObjectN.bin`，编号不连续、跨页复用，部分不是 OLE2 复合文件（可能是链接对象/包/图片）。

不要假设：
- 第 3 页 = `oleObject3.bin`；
- 第 5 页 = `oleObject5.bin`。

必须解析：

```text
ppt/slides/slideN.xml
  <p:oleObj r:id="rIdX" ...>
ppt/slides/_rels/slideN.xml.rels
  rIdX -> ../embeddings/oleObjectM.bin
```

`ppt_ole_ops.list_ole_objects()` 就是做这件事。

## 3. `op.open()` 返回 False 不代表没打开

对从 OLE 拆出来的旧工程，`op.open()` 经常返回 `False`，但 `op.graph_list()` 已经有图。判断是否成功要看 `graph_list()`，不要只看返回值。

## 4. 读图里的数据

图里的数据是引用，不是副本。`pl.obj.GetRange()` 返回 `[Graph1]1!Plot(1)`，`DataRange.GetData()` 对它不可靠。

可靠写法：

```labtalk
range -wx rx = 1;
range -wy ry = 1;
```

然后：

```python
n = op.lt_int('rx.GetSize()')
x = [op.lt_float(f'rx[{i}]') for i in range(1, n+1)]
y = [op.lt_float(f'ry[{i}]') for i in range(1, n+1)]
```

`op.lt_float()` 能求 LabTalk 表达式，`op.get_lt_str()` 只能读 string 变量。

## 5. 轴单位转换

- 改标题：`ax.title = ...`（originpro 的 Axis 对象）；
- 改刻度显示：`layer.x.label.formula$ = "x+273.15"`；
- 不要改数据和轴范围，否则 PPT 上已有的手绘标注会错位；
- 富文本标题要保留 `\pNNN(...)`、`\f:Times New Roman(...)` 等标签；
- `layer.x.title$`、`layer.x.label.text$` 在 2025b 里不可靠，`ax.title` 最稳。

## 6. 图例被拟合线污染

加入拟合 plot 后，Origin 的 `Legend` 自动更新，出现 `GSD`、`ConvD`、`EnvLow` 等条目。

- `legend.text$` 可以改，但 `legendupdate` 或保存/重开可能又重建；
- 最稳：`gl.remove_label('Legend')`，然后用 `gl.add_label()` 加静态标签；
- 静态标签不会随拟合 plot 变化；
- 标签字体用 `label.set_float('fontsize', 20)`；颜色用 `label.color = (r,g,b)`。

## 7. 线型/线宽

- `set %C -l 2` 是**线型**（虚线），不是线宽；会让拟合线变虚线；
- 线宽用 `set %C -w 2`；
- 颜色用 `set %C -c 4` 或 `plot.color = (0,0,255)`；
- 不要在已经写过 `-l 2` 的图上继续调，最好从干净工程重新加拟合线。

## 8. 填充包络

- `Plot.set_fill_area(above, 9, below)` 的 `above/below` 是 Origin 颜色索引，不是 RGB；
- 可靠写法：
  ```labtalk
  set %C -pfv 8;   // fill to next data plot
  set %C -pf 1;    // enable
  set %C -pfn 1;   // common X area
  set %C -cf 5;    // fill color index
  ```
- `color(200,230,255)` 的值是 `33547976`，`-cf 33547976` 可以得到浅蓝色；
- 包络两条边界线要保留，否则填充可能消失；
- `plot.transparency = 65` 可能让填充整体消失，慎用。

## 9. WPS vs PowerPoint

- WPS 常走 VML `<mc:Choice>` 分支，使用 WMF 预览；
- 现代 Office 走 `<mc:Fallback>` 的 EMF 图片；
- 用 Origin 导出的 WMF 替换 VML 的 WMF 时，WPS 可能渲染成整页斜线；
- 最稳：删除 `<mc:Choice>...</mc:Choice>`，强制 WPS/Office 都用 EMF 回退图；
- 验证必须用 WPS；本机 PowerPoint 无授权，而且它的 OLE 分支行为不同。

## 10. 重建 OLE 复合文件的时间戳

用 `pythoncom.StgCreateDocfile` 重建 OLE 时，根目录项会写入当前时间戳。原文件中这些时间戳可能是 0。PowerPoint 可能因此报：

```text
OLEFormat: Object does not exist
```

解决：
- 把根目录项/流目录项的 creation/modified time 8 字节归零；
- 或者更推荐：不要手工重建 OLE，直接用 Origin copy → WPS PasteSpecial 生成工厂 OLE。

## 11. `op.exit()` 会断开 OLE 连接

在 PPT 激活的 OLE 服务器里修改后：
- `op.save()` 会报 `must supplied a path if project not saved yet`；
- `op.save(path)` 会另存为独立 OPJU，不能写回嵌入对象；
- `op.exit()` 会让 PowerPoint/WPS 的 `ActivePresentation` 失效。

所以不要试图“在嵌入工程里 op.save 再 op.exit”，直接用工厂 OLE 替换。

## 12. 幻灯片布局：图被拖走

判断形状是否被移动：
- 解析 `slideN.xml` 的 `a:off`/`a:ext`，与历史快照比较；
- 或 WPS COM 读取 `Shape.Left/Top/Width/Height`（单位 pt）。

实际遇到的例子：
- slide 5 的上半图 `对象 88` 在 L=47, T=17.75（左边，被 TEM 图盖住）；
- 它的蓝/黑曲线 `形状 89/90` 在 L=512–596（右边）；
- 移动 `对象 88` 到 L≈498, T≈-30, W≈443, H≈270 后，曲线、I/II、虚线重新对齐。

## 13. 拟合模型与数据分组

对硬度图：
- 红色 GSD Cu 是一组，logistic R²≈0.993；
- 蓝色 "Conventional Cu" 实际是 **三组不同数据**（空心圆/方块/三角），单条 logistic R²≈0.84；
- A 方案：三组各一条 logistic + 包络，最忠实；
- 包络较宽是数据本身差异，不是拟合失败。

其他：
- d_G / d_D：指数增长；
- f_domain / Zener pinning：高斯；
- 驱动压力黑色：二次多项式 R²≈0.99（指数只有 0.88）；
- DSC：原生线，不需要拟合。

## 14. 永远不要直接改原文件

- 先 `cp` 一份；
- 所有替换/移动/删除都在副本上；
- 验证通过后再交付新文件名，原文件保持不动。

## 15. 验收清单

- [ ] WPS 双击 OLE → Origin 出现 `Graph1`
- [ ] originpro attach 后 `len(gl.plot_list())` = 原 plot 数 + 拟合/包络数
- [ ] X 轴标题含 `(K)`，公式 `x+273.15`
- [ ] 幻灯片渲染里手绘拟合线已删除，原生拟合线可见
- [ ] I/II、虚线、EXO、C/D、TEM 图片、文字位置未变
- [ ] 非目标页像素/结构未变化
- [ ] 原 PPT 未被修改

## 16. 图在框里只占 ~83%：问题不在 OLE 内容

现象：替换后图能双击、内容也对，但在原图框里明显变小，四周一圈白边。

原因有两层：

1. 工厂 OLE 的预览是整页（含页边距），绘图区只占页面的一部分；
2. 预览的长宽比与幻灯片图框不一致，WPS 按长宽比适配后进一步缩小。

先量，不要猜：

```bash
python scripts/ole_inspect.py frames deck.pptx    # 每个 OLE 的 imgW/imgH 与 frame aspect
python scripts/ole_inspect.py factory factory.pptx
```

判断：

- `delta%` 很小（<3%）：尺寸问题基本只剩页边距；
- `delta%` 很大（例如 -35%）：先别硬塞，多半是工厂里混入了页外对象（见 #17）。

## 17. `SetExportClip(True)` 不是“裁掉页边距”

它实际会把**页外对象**一起纳入导出，并改变预览的物理尺寸。对于没有页外垃圾的图，看起来“填满了”；对于有残留对象的图，会把长宽比拉爆。

典型事故：一张 DSC 图页面右侧躺着一个早先操作留下的 `Legend`（文字类似 “Subtracted from B”）。普通导出看不到它；`ExportClip` 一开，它被完整导出，工厂原生尺寸从 1.3:1 变成 2.0:1，替换后图被压成一条。

排查与处理：

```python
gp = op.graph_list()[0]; gp.activate()
gos = gp[0].obj.GraphObjects
for go in gos:
    print(go.Name, go.Left, go.Top, go.Width, go.Height)  # 1/1000 inch
gos['Legend'].Destroy()     # 没有 .Delete()
```

判断“页外”的参照：页面宽高 `gp.obj.Width/Height`（单位 inch）。`Left` 远大于页面宽度的对象就是嫌疑对象。

## 18. 同一个 EMF 预览会被多个 OLE 共享

同一张图在 PPT 里出现两次（例如概览页和详情页）时，两个 `oleObjectN.bin` 可能引用同一个 `ppt/media/imageNN.emf`。

后果：替换时“最后一个写入的媒体”会同时作用于所有引用它的对象。如果两个对象需要不同的预览（比如一个图框高、一个图框矮），必须：

- 给它们分配同一个工厂；或
- 改完后再单独处理其中一个页面的媒体引用。

不要假设“每个 OLE 都有独立的预览图”。

## 19. 轴标题压住 K 刻度

把 `ax.title` 改成 `Temperature (K)` 后，某些图会出现标题与刻度重叠：Celsius 刻度是 `0/100/200`，换成长度翻倍的 `273.15/373.15/...` 后，标题的固定位置就不够了。

可用的调整（按优先级）：

1. 下移 X 轴标题对象：
   ```python
   xb = gp[0].obj.GraphObjects['XB']   # X Bottom = X 轴标题
   xb.Top += 60                        # 单位 1/1000 inch
   ```
2. 同时略缩小刻度字号：`gp[0].obj.SetNumProp('x.label.pt', 14)`。

注意：**移动标题会增加预览的高度**。如果目标图框很矮（例如概览页里的小图），标题会被顶到图框外、压到相邻图。必须逐页渲染确认；必要时为同一张图在“大图框页”和“小图框页”分别生成两个工厂。

## 20. PPT 上叠了一层预览图（重影）

WPS 有时会同时显示 OLE 对象和 `p:oleObj` 内部的 `<p:pic>` 预览图，看起来像半透明重影。

处理：

```python
from ole_inspect import strip_ole_previews
strip_ole_previews('deck.pptx', 'deck_no_preview.pptx')
```

- 只删 `p:oleObj` 内部的 `p:pic`；页面上的真实图片（SEM 照片等）不能动；
- 删完 OLE 仍从嵌入文档渲染，双击编辑正常；
- 验证：`p:pic inside oleObj = 0`，并抽 2~3 页渲染 + 双击确认。

## 21. 大文件不要反复整页渲染

一份 100 MB+ 的 PPT，每次 WPS 打开 + 导出都要几十秒，还会产生大量图片进入上下文。推荐节奏：

1. 在工厂层用 `ole_inspect.py` 判断（纯 zip 操作，秒级）；
2. 用 `ole_inspect.py preview` 把 EMF 转 PNG 直接看；
3. 方案确定后，只做一轮整页渲染验收；
4. 读图时裁剪关键区域，不要每次都读整页大图。

这套顺序能把一次 OLE 批处理的工具调用量压到原来的几分之一。
