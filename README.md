# 🎬 VideoHub

> 极简 · 强大 · 天生为探索而生  
> 用苹果式的直觉，管理你珍藏的每一帧记忆。

本地视频文件的智能管理器。  
扫一眼、搜一秒、点一下，全部搞定。

---

## ✨ 核心体验

- **即时搜索** — 输入即筛选，毫秒级响应
- **智能解析** — 自动读取分辨率、时长、大小、修改时间，甚至 Windows 文件标题与隐藏链接
- **多线程扫描** — 千部影片也能优雅加载
- **内置 Web 服务** — 一键启动，浏览器即客户端
- **响应式玻璃态 UI** — 毛玻璃 + 柔和阴影，宛如 macOS 原生
- **一键跳转** — 有链接直达来源，无链接自动 Bing 搜索文件名
- **导出清单** — 一键生成 `log.txt`

---

## 🚀 快速开始

```bash
git clone https://github.com/d1irgger/VideoHub.git
cd VideoHub/VideoHubWeb版本VideoHubWeb\ version
pip install opencv-python pywin32
python kkp1.py
浏览器打开 http://localhost:8000，选择文件夹，即刻探索。
仅支持 Windows（依赖 pywin32 读取文件标题）。

📸 预览

实际界面比文字更惊艳 — 运行即见。

📂 仓库结构





















目录说明VideoHubWeb版本...推荐 — 当前 Web 版主程序VideoHub早期原生UI版本...早期 Tkinter 原生界面web实现论证...技术验证与实验代码

🛠️ 技术栈
Python · OpenCV · 多线程 · 内置 HTTP Server · 原生玻璃态前端

📬 反馈
由 @d1irgger 设计开发。

欢迎 Issue / PR。
如果喜欢，给一颗 Star — 这是最大的鼓励。
