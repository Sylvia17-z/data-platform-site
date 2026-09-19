# 数据统计平台 · 云端部署说明（GitHub Pages / Gitee Pages）

> 目的：实现「人不在电脑边、电脑关机、不在同一 WiFi」也能随时用手机/电脑访问。
> 原理：平台是**纯前端单文件**（index.html 电脑版 / mobile.html 手机版），不需要服务器程序，
> 任何静态托管都能直接跑。GitHub Pages 免费、永不休眠。

---

## 一、先搞清楚两个链接（部署完成后）

| 端 | 访问地址 |
|---|---|
| 电脑版 | `https://你的用户名.github.io/stats-platform/` |
| 手机版 | `https://你的用户名.github.io/stats-platform/mobile.html` |

> 手机建议用浏览器「添加到主屏幕」，当 App 一样用。
> 电脑在单位/家里访问同一个地址即可，与当前这台电脑是否开机**完全无关**。

---

## 二、GitHub 部署步骤（主推）

### 第 1 步：注册 GitHub
打开 https://github.com ，没有账号先注册（邮箱验证），已有则登录。

### 第 2 步：新建空仓库
右上角 `+` → **New repository**
- Repository name：`stats-platform`
- 选 **Public**（免费 Pages 必须公开；仓库里只放网页文件，不放数据）
- **不要勾选** "Add a README file"（保持空仓库，后面一键脚本才不会冲突）
- 点 **Create repository**

### 第 3 步：一键发布（推荐，之后每次改版都只用这一步）

**本机需已安装 Git**（cmd 里 `git --version` 能出版本号即可；没有就去 https://git-scm.com 装，全程下一步）。

然后**双击**项目文件夹里的：`发布到GitHub.bat`
- 第一次运行：按提示粘贴仓库地址 `https://github.com/你的用户名/stats-platform.git` 回车
- 会自动弹出 GitHub 登录窗口（Git Credential Manager）→ 浏览器里登录并授权
- 看到 `[完成]` 即推送成功

> 之后每次我帮你改完代码、重新生成 index.html / mobile.html 后，
> **你只需再双击一次这个 bat**，约 1 分钟线上即更新 —— 全程不用我再发布。

### 第 4 步：开启 Pages（只需一次）
1. 仓库页 → **Settings** → 左侧 **Pages**
2. Build and deployment → **Source**: 选 `Deploy from a branch`
3. **Branch**: `main` / `(root)` → **Save**
4. 等 1~2 分钟，页面顶部出现 `Your site is live at https://你的用户名.github.io/stats-platform/`

### 第 5 步：验证
- 手机浏览器打开 `https://你的用户名.github.io/stats-platform/mobile.html`
- 导入一份真实 Excel 试跑（预约表/订单表均可）
- 能正常统计即部署成功

---

## 三、不用 Git 的备选上传方式（网页直接传）

若本机没装 Git，也可以每次用网页上传：
1. 仓库页 → **Add file** → **Upload files**
2. 把 `index.html`、`mobile.html` 两个文件拖进去 → **Commit changes**
3. 之后每次更新 = 重新上传这两个文件覆盖（会提示 Replace）

缺点：每次改版都要手动拖文件；有 Git 的话用 bat 一次搞定。

---

## 四、Gitee 备选（若 GitHub 打不开）

1. https://gitee.com 注册并完成**实名认证**
2. 新建仓库 `stats-platform`（公开）→ 网页上传两个文件，或用 bat 改填 Gitee 仓库地址推送
3. 仓库页 → 服务 → **Gitee Pages** → 启动（首次可能有人工审核）
4. ⚠️ 与 GitHub 不同：**每次更新代码后，需回 Gitee Pages 页面手动点「更新」按钮**才会同步

---

## 五、重要注意事项

| 事项 | 说明 |
|---|---|
| 仓库必须 Public | 免费 Pages 不发布私有仓库；仓库里只有 2 个网页文件，无敏感数据 |
| **绝不把 Excel 数据源传上去** | 仓库是公开的，客户数据/表格一律不要上传，数据始终在你自己浏览器里 |
| 云端是全新域名 | 浏览器本地存储按域名隔离 → **首次打开需重新导入一次数据**；同一浏览器之后再打开会保留 |
| 双击 file:// 电脑版照常用 | 云端只是新增通道，不替代本地使用习惯 |

---

## 六、以后每次「改版更新」的固定流程

1. 告诉我改什么 → 我修改并重新生成 `index.html` / `mobile.html`（本地先自测）
2. 你**双击 `发布到GitHub.bat`**（或网页重新上传两个文件）
3. 等约 1 分钟，手机 Ctrl+F5 / 重新打开即新版
