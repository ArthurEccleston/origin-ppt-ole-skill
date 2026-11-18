# 排障提示词（可直接粘贴给任意 AI）

> 用途：当你接手一个“PPT 里的 Origin 图不对”的任务时，把下面这段连同你的现象一起发给 AI，可以避免它重复最常见的错误。

---

我在 Windows 上处理 PowerPoint/WPS 里嵌入的 Origin 图（OLE 对象）。请按以下事实工作，不要凭直觉试错：

**已知事实**

1. PPT 里的 Origin 图是“嵌入 Origin 文档”，不是独立 `.opju`。把独立 OPJU 写回 OLE 的 `Contents` 流会让双击失效。
2. 合法的替换方式是：Origin 中 `gp.copy_page('OLE')` → WPS 中 `Shapes.PasteSpecial(10)` 生成一个“工厂 PPTX”，取出其中的 `ppt/embeddings/oleObject1.bin` 和 `ppt/media/image1.emf`，再替换目标 PPT。
3. `oleObjectN` 编号和幻灯片顺序无关，必须解析 `ppt/slides/_rels/slideN.xml.rels` 的 `rId → oleObjectN.bin` 映射。
4. 验证以 WPS 为准（WPS 走 VML/EMF 分支，PowerPoint 行为不同）。但“双击能开”不等于“显示正确”。
5. `gp.obj.SetExportClip(True)` 的真实语义是“把页外对象也导出”，不是“裁掉页边距”。它会改变预览物理尺寸，页外残留的 Legend 会把长宽比拉爆。
6. `p:oleObj` 里的 `<p:pic>` 是预览图。WPS 可能同时显示 OLE 和这张图，造成重影；删掉 OLE 内部的 `p:pic` 不影响双击编辑。
7. 同一张图在多个页面出现时，可能共享同一个 `ppt/media/imageNN.emf`；替换一个会波及全部引用者。
8. 幻灯片图框的 `a:ext` 与 `p:oleObj` 的 `imgW/imgH` 长宽比不一致时，图会被 letterbox 或溢出。先量长宽比，再决定用哪个工厂。

**工作方式**

- 不要每次修改都渲染整份 PPT。先在包内用脚本检查：OLE 清单、imgW/imgH、frame ext、EMF 的 `rclFrame` 物理尺寸；EMF 可以直接用 Pillow 打开成 PNG 查看。
- 方案确定后，只做一轮整页渲染验收；读图时裁剪关键区域，不要反复读整页大图。
- 永远在副本上操作，原文件不动。
- 报告时区分“已修复”和“仍然存在”，不要用“应该没问题”代替验证。

**请先给出的信息**

1. 目标 PPT 的 OLE 清单（slide / rId / embedding / fallback media / imgW / imgH / frame ext）。
2. 每个工厂 OLE 的 imgW/imgH 与 EMF 物理长宽比。
3. 你判断的根因，以及对应的可复现证据（尺寸、长宽比、渲染图）。

我的现象是：**（在这里写：例如“第 3 页图只占框的 83%”“图上叠了一层图”“双击空白”）**

---

## 给 Claude Code 的短版

如果当前环境已经装了 `origin-ppt-ole-skill`，直接说：

> 用 origin-ppt-ole-skill 处理这个 PPT 里的 Origin OLE；先跑 `ole_inspect.py frames/factory`，不要整页渲染，最后只做一轮验收。
