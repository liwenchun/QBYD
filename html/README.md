---
AIGC:
    Label: "1"
    ContentProducer: 001191440300708461136T1XGW3
    ProduceID: 745d21d7788b4659dd78c7f722469398_d4f2f51bd9bc11f19232525400248c00
    ReservedCode1: LZvPDWpB/J1+pK7aU2ImEFoVlOyQoc/+ocNBBDvEJXAMQ0Z0Y3ke5OtZqvYsxOXISuqddzVq7+Rc9XRGGTanIwUcZw5J8M7847qKz1oMRjyIwZVSXUV2qbEKbPE/dKiVFHj14K4wv8aickXilLig349Sgg3cpUXNdMzc+OjuqXzMXlYNM5i2Jvm0Pcs=
    ContentPropagator: 001191440300708461136T1XGW3
    PropagateID: 745d21d7788b4659dd78c7f722469398_d4f2f51bd9bc11f19232525400248c00
    ReservedCode2: LZvPDWpB/J1+pK7aU2ImEFoVlOyQoc/+ocNBBDvEJXAMQ0Z0Y3ke5OtZqvYsxOXISuqddzVq7+Rc9XRGGTanIwUcZw5J8M7847qKz1oMRjyIwZVSXUV2qbEKbPE/dKiVFHj14K4wv8aickXilLig349Sgg3cpUXNdMzc+OjuqXzMXlYNM5i2Jvm0Pcs=
---

# 秦巴药都 · 中药材全产业链数据平台（纯静态）

无后端、无数据库、无第三方服务，全部由浏览器直接读取本地 xlsx 完成。可直接部署到 GitHub Pages / 内网静态服务器，也可在本机用 Python 自带的 http.server 运行。

## 一、目录结构

```
output\
├─ start-login.bat          # 一键启动本地服务并打开登录页（服务根 = output）
├─ html\
│  ├─ index.html            # 官网首页（对外宣传页）
│  ├─ login.html            # 登录页（读取 data/usersData.xlsx 校验账号）
│  ├─ home.html             # 登录后主页面（数据平台）
│  ├─ README.md             # 本说明
│  └─ js\xlsx.full.min.js   # SheetJS，本地引入，不依赖 CDN
├─ data\
│  ├─ usersData.xlsx        # 账号 + 权限矩阵（用户 / 用户权限 / 说明 三张表）
│  ├─ usersData.js          # 账号与权限的本地快照（读取 xlsx 失败时的兜底）
│  ├─ homeData.xlsx         # 业务数据：种植 / 采收 / 入库 / 生产 / 出库 / 销售 六张明细表
│  └─ homeData.js           # 业务数据的本地快照（双击打开 html 时使用）
└─ images\                  # 首页用图
```

## 二、本地启动

**推荐**：双击 `output\start-login.bat`，浏览器会打开 `http://localhost:8000/html/login.html`，读取的是 data 目录里的真实 xlsx。

**也可以直接双击** `output\html\login.html`（`file://` 协议）：登录、首页、数据表、增删改、导出都能用，但此时浏览器不允许网页读取同目录的 xlsx，数据来自快照文件 `data\homeData.js`；若要读取最新 Excel，在页面顶部提示条点「手动选择 homeData.xlsx」，或改用上面的 start-login.bat。
注意：快照不会自动跟随 Excel 变化，改完 `homeData.xlsx` 后建议用 start-login.bat 打开。

示例账号：`user / 123456`（一般用户）、`manager / 123456`（管理用户）、`admin / 123456`（超级管理用户）。

## 三、登录与会话规则

- 登录成功后跳转 `home.html?user=<用户名>&ts=<登录时间戳(毫秒)>`，**密码不进地址栏、不进 sessionStorage**。
- `home.html` 只接受这两种来源的会话：地址栏 `user`+`ts`、或 `sessionStorage.qbd.auth`。
- **会话有效期 5 分钟**：页面顶栏实时显示倒计时；超时（含刷新时校验）自动跳回 `login.html?expired=1` 并提示重新登录。
- 缺少会话参数直接访问 `home.html` → 跳回 `login.html?reason=auth`。
- 想改成跳官网首页，把 `home.html` 里 `CONFIG.loginPage` 替换为 `CONFIG.indexPage` 的值即可（`./index.html`）。

## 四、home.html 页面规则

| 区域 | 说明 |
| --- | --- |
| 顶栏 | 用户名、角色、登录时间、会话倒计时、退出 |
| 左侧菜单 | 由权限表动态生成，只列出当前账号可见的表；另有一项「首页」 |
| 首页 | 品牌图欢迎页 + 各表记录数概览卡片；**不直接展示表格**，点菜单或卡片才进入数据表 |
| 数据视图 | 表头（支持 Excel 多级合并表头）、分页、搜索、增删改 |

- **分页**：默认每页 10 条，可切 20 / 50；搜索框对当前表全列模糊过滤，过滤后再分页。
- **表头**：自动识别前 1~4 行表头（含合并单元格），滚动时表头吸顶，横向滚动保持对齐。
- **空表**：`homeData.xlsx` 里只有表名没有内容的表，显示空态提示；有写权限时可点「新增行」建立默认列（序号 / 品种 / 日期 / 数量 / 经办人 / 备注）后录入。想要自定义列名，先在 Excel 里写好表头再刷新页面。
- **编辑是内存态**：增、删、改只改浏览器内存，**不会自动写回文件**。点「导出 xlsx」才会把当前表连同其余未改动的工作表一起导出为同名文件（`homeData.xlsx`），需要手动覆盖 `data\` 目录中的原文件。点「撤销编辑」可放弃本次改动，恢复文件原始数据。
- **导出时表头会扁平化**：Excel 的多级合并表头在导出文件中会变成「父级-子级」的单行列名（如 `中药材基地种植记录信息-栽种阶段-栽苗浇水(定根)`），这是为了避免合并单元格导致再次读取错位。
- **用户明细**：对应 `usersData.xlsx` 的「用户」表，一般用户不可见；管理用户可见可改但密码列显示为掩码、只能重置不能查看；仅超级管理用户可以查看明文密码并导出。

## 五、权限配置（改 Excel 即可，无需改代码）

`usersData.xlsx` →「用户权限」表：

| 列 | 含义 |
| --- | --- |
| 表名 | 与 `homeData.xlsx` 的工作表名一致；「用户明细」代表 usersData 的「用户」表 |
| 一般用户 / 管理用户 / 超级管理用户 | 1 = 该角色可见此表，0 = 不可见 |
| 可删除 | 1 = 管理用户可删除数据；0 = 只能增改不能删。一般用户恒为只读，超级管理用户不受此列限制 |

角色最终落到的操作级别：

| 角色 | 级别 | 能做什么 |
| --- | --- | --- |
| 一般用户 | r | 只看，不能增删改；只能看权限表里一般用户=1 的表 |
| 管理用户 | rw / rwd | 可见表可增可改，可删除列=1 时才出现「删除」按钮 |
| 超级管理用户 | rwd | 全部表可见，全部可增删改 |

「用户」表：`用户名 / 用户密码 / 权限 / 姓名 / 备注`。用户名匹配不区分大小写；密码为明文存储，仅作演示，请勿用于真实生产系统。

改完 `usersData.xlsx` 后，建议同步更新 `data\usersData.js` 快照（结构见该文件），否则 xlsx 读取失败时的兜底数据仍是旧的。

## 六、部署到 GitHub Pages

1. 新建 Public 仓库，把 `html`、`data`、`images` 三个目录整体上传（保持相对位置）。
2. Settings → Pages → Source 选 `Deploy from a branch`，Branch 选 `main` / `/(root)`，保存。
3. 访问 `https://<用户名>.github.io/<仓库名>/html/login.html`。
4. 修改 `data\*.xlsx` 后重新上传，Ctrl+F5 强刷即可看到新数据（`html/js` 与 xlsx 均已用 `cache: no-store` 拉取）。

## 七、常见问题

| 现象 | 原因与处理 |
| --- | --- |
| 双击 login.html，登录后首页只显示旧数据 | `file://` 下读不到 xlsx，用 `data\homeData.js` 快照。在页面顶部提示条点「手动选择 homeData.xlsx」读当前文件，或用 `start-login.bat` |
| 双击打开后想更新快照 | 打开 `html\生成数据快照.html`，选 `data\homeData.xlsx` → 生成并下载 `homeData.js` → 覆盖 `data\homeData.js` |
| 提示「数据加载失败」 | 确认 `data\homeData.xlsx`、`data\usersData.xlsx` 存在且与服务同源；双击模式还需 `data\homeData.js`、`data\usersData.js` 存在 |
| 提示「账号不存在或已被删除」 | 用户名不在 usersData 的「用户」表中，注意大小写不敏感但拼写要对 |
| 登录后 5 分钟被踢回登录页 | 设计如此（会话有效期 5 分钟），改 `home.html` 里 `CONFIG.ttlMs` 可调整 |
| 导出的文件打开后列名变长 | 多级表头被扁平化为「父-子」，属预期行为 |
| 改了数据刷新却没了 | 编辑只在内存，未导出的改动刷新即丢失 |

*（内容由 AI 生成，仅供参考）*
