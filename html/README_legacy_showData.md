---
AIGC:
    Label: "1"
    ContentProducer: 001191440300708461136T1XGW3
    ProduceID: 745d21d7788b4659dd78c7f722469398_bb8e045ad8f511f19280525400393706
    ReservedCode1: epp3s1lv54DRXodUg7XUAWI5cHhNKNirBhWrotTy+E7qA6VL+c4tjhbE4zVKhPED7xXy/bVXWV9DSyECoalH4gwfM/XlCM/OXgkxhZ3n5YcTtzrQi4qh0Zko2frCNBUCd3IfLbMpLZjoOeNj0W8HU5TvuqnGfRAl5GunEHobRIn7j1l60YtmiUIwjrQ=
    ContentPropagator: 001191440300708461136T1XGW3
    PropagateID: 745d21d7788b4659dd78c7f722469398_bb8e045ad8f511f19280525400393706
    ReservedCode2: epp3s1lv54DRXodUg7XUAWI5cHhNKNirBhWrotTy+E7qA6VL+c4tjhbE4zVKhPED7xXy/bVXWV9DSyECoalH4gwfM/XlCM/OXgkxhZ3n5YcTtzrQi4qh0Zko2frCNBUCd3IfLbMpLZjoOeNj0W8HU5TvuqnGfRAl5GunEHobRIn7j1l60YtmiUIwjrQ=
---



# Excel 网页展示（GitHub Pages）

一个纯静态页面，读取**同目录下的 `data.xlsx`**，把表格数据渲染到网页上。不需要后端、不需要数据库。

## 文件说明

| 文件 | 作用 |
| --- | --- |
| `showData.html` | 页面本体（含全部样式和逻辑，依赖 CDN 上的 SheetJS 解析 Excel） |
| `data.xlsx` | 你的数据文件，替换成自己的即可（文件名必须叫 `data.xlsx`） |
| `README.md` | 本说明 |

## 数据路径规则（重要）

页面读取表格用的是**相对路径** `./data.xlsx`，含义是「与 `showData.html` 同一目录下的 data.xlsx」。

- 部署在 `https://用户名.github.io/仓库名/` 时，实际请求地址是 `https://用户名.github.io/仓库名/data.xlsx`
- **不要**写成绝对路径（如 `/data.xlsx` 或 `https://用户名.github.io/data.xlsx`），否则仓库部署在子目录下会 404
- 表格必须与 `showData.html` 放在**同一层目录**，不要放进子文件夹
- 代码中用 `new URL('./data.xlsx', document.baseURI)` 解析，换目录、换仓库名都不用改代码

## 部署到 GitHub Pages

1. 登录 GitHub，点右上角 **+ → New repository**，仓库名例如 `excel-viewer`，可见性选 **Public**（免费版 Pages 需要公开仓库），创建。
2. 进入仓库，点 **Add file → Upload files**，把 `showData.html`、`data.xlsx`、`README.md` 三个文件拖进去，**Commit changes**。
3. 仓库页 **Settings → 左侧 Pages**：
   - Source 选 **Deploy from a branch**
   - Branch 选 **main**，目录选 **/ (root)**，点 **Save**
4. 等 1~2 分钟，刷新该页面，顶部会出现网址：
   `https://你的用户名.github.io/excel-viewer/`
5. 打开时**必须带上文件名**：
   `https://你的用户名.github.io/excel-viewer/showData.html`

> 首次生效通常 1~3 分钟；之后每次修改 `data.xlsx` 并提交，刷新网页就是最新数据（建议 Ctrl+F5 强刷避免缓存）。

### 关于访问地址（文件名不是 index.html 的影响）

GitHub Pages 只把 `index.html` 当作目录的默认首页。页面改名成 `showData.html` 后，直接访问仓库根地址会显示 404。三种处理方式：

| 方式 | 做法 | 效果 |
| --- | --- | --- |
| 访问完整路径（最简单） | 用 `.../仓库名/showData.html` | 立即可用，不需额外文件 |
| 加跳转页 | 在仓库放一个 `index.html`，内容只有一行跳转 | 访问根地址自动跳到 showData.html |
| 改回 index.html | 文件名改回 `index.html` | 根地址直接打开 |

## 本地预览

直接双击 `showData.html` 是不行的——浏览器对 `file://` 协议禁止 fetch 本地文件。请在文件所在目录执行：

```bash
python -m http.server 8000
```

然后浏览器访问 `http://localhost:8000/showData.html`，或双击 `start-local.bat` 自动完成。

## 页面功能

- 自动加载同目录 `data.xlsx`
- 多工作表自动生成切换标签
- 顶部搜索框可跨列模糊过滤
- 点击表头排序（数字按数值排、文本按拼音排，再点切换升降序）
- 支持 xlsx / xls / csv；也可点「选择本地文件」或直接拖文件进来临时查看（数据只在本机解析，不上传）

## 换成自己的数据

两种方式：

1. **保持文件名**：把自己的表格改名为 `data.xlsx`，覆盖原文件后提交。
2. **改代码里的路径**：在 `showData.html` 中搜索 `var DEFAULT_FILE = './data.xlsx';`，改成你要的相对路径（例如 `'./销售数据.xlsx'`）。改的时候保留前面的 `./`，注意文件名大小写要与实际文件完全一致。

注意：中文文件名在 GitHub Pages 上一般可用，但建议用英文或数字命名更稳妥。

## 适配自己的表头

页面把 Excel **第一行当作表头**，从第二行开始作为数据行。如果你的表格前面有标题行、合并单元格等，请先整理成「第一行是列名」的规范格式。

## 数据量说明

前端渲染，实测几千行无压力；超过 2 万行建议分页或改用接口。如果需要分页、图表、导出等功能，可以在此基础上继续加。

## 常见问题

- **页面显示"没有读取到 data.xlsx"**：确认文件名大小写一致、且和 `showData.html` 在同一目录（不是放进子文件夹）。
- **改了 Excel 但网页没变**：GitHub Pages 有缓存，Ctrl+F5 强制刷新，或等 1~2 分钟。
- **表格是空白的**：用 Excel 打开确认数据在当前可见的工作表里，没有放在隐藏表。
*（内容由AI生成，仅供参考）*
*（内容由AI生成，仅供参考）*
