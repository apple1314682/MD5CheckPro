# MD5 校验工具 Pro

macOS 风格的 Windows 桌面 MD5 校验工具：拖入即算，结果自动写入 Excel。

![icon](icon.png)

**V2.4：文件卡片鼠标悬停显示 ▲▼，点击上下微调顺序，彻底告别误拖拽。**

**V2.3：文件卡片支持拖拽上下排序（复制 / 导出TXT 顺序跟随）；启动再提速——体积减半、WebView2 用户目录持久化、单实例互斥。**

**V2.2：新增勾选删除——勾选文件后「删除选中」可同时移除列表卡片与 Excel 对应行；「清空」按钮同步清空 Excel 全部记录（含历史）。V2.1：单文件 exe 发布，免 Python 环境，双击即用；启动画面 + 窗口就绪后才显示，告别黑屏等待。**

## 下载使用

三种方式任选：

1. **直接下载 exe**：到本仓库 [Releases](../../releases) 页面下载 `MD5校验工具.exe`（约 15 MB 单文件，免安装、免 Python，可拷到任意目录运行）
2. **自己打包**：克隆仓库后按下方「从源码运行 / 打包」一节执行一条 pyinstaller 命令
3. **源码运行**：`pip install pywebview openpyxl` 后 `python MD5校验工具.py`

- 记录表 `MD5校验记录.xlsx` 自动生成在 exe 所在目录
- 系统需有 WebView2 运行时（Win10/11 一般自带）；缺失时程序会弹窗提示安装地址

## 功能

- **拖入即算**：文件（支持多选）直接拖进窗口自动计算；拖入文件夹自动递归扫描批量计算
- **选择文件 / 选择文件夹**：按钮入口，与拖拽等效
- **校验比对**：粘贴已知 MD5（32 位），算完自动显示 ✅一致 / ❌不一致
- **Excel 自动记录**：每算完一个文件立刻把「文件名 + MD5」等追加写入 exe 目录的 `MD5校验记录.xlsx`（含大小、创建/修改时间、耗时、文件路径），持续累积；表格被 Excel/WPS 占用时自动另存新表
- **复制 / 导出 TXT**：一键复制全部结果，或另存 txt 报告
- **上下微调排序（V2.4）**：鼠标移到文件卡片上，右侧浮现 ▲▼ 按钮，点击上下移动；「复制 / 导出TXT」的报告顺序跟随列表
- **勾选删除（V2.2）**：勾选任意文件卡片，点「删除选中」，列表卡片与 Excel 中对应行一起删除；排队/计算中的任务也会安全取消
- **停止 / 清空**：中断计算、清空列表；「清空」同时清空 Excel 记录表的全部数据行（包括之前累积的历史记录）

## 界面

macOS 深色风格：无边框窗口 + 红绿灯按钮、毛玻璃卡片、实时进度条、逐文件耗时展示。启动时先显示应用图标画面，主界面就绪后直接呈现（不再黑屏等待）。

## 从源码运行 / 打包

环境要求：Windows 10 / 11、Python 3.10+、依赖 `pywebview`、`openpyxl`。

```bat
pip install pywebview openpyxl
python MD5校验工具.py
```

打包单文件 exe（命令已写入主程序文件头注释）：

```bat
pip install pyinstaller
pyinstaller --onefile --noconsole --icon app.ico --name "MD5校验工具" ^
    --add-data "ui.html;." --add-data "app.ico;." --add-data "bottle.py;." ^
    --collect-all webview --collect-all pythonnet --collect-all clr_loader ^
    --exclude-module bottle --splash icon.png MD5校验工具.py
```

> bottle.py 是从 site-packages 复制到本目录的完整副本：Python 3.10 下 PyInstaller 静态扫描大文件 bottle.py 会崩溃，故排除分析、以数据文件形式随包，运行时正常导入。

> 注意：pywebview 需与本机 WebView2 运行时版本匹配。开发机用的 pywebview 6.2.1 需要较新的运行时（2025 年后更新的 Win10/11 均可）。

## 目录结构

```
MD5校验工具Pro/
├── MD5校验工具.py        # 主程序（MD5 计算、Excel 记录、窗口控制）
├── ui.html               # 界面（单文件 HTML/CSS/JS）
├── app.ico / icon.png    # 应用图标（icon.png 兼作打包启动画面）
├── bottle.py             # 随包分发的 bottle 运行时（打包用，见上方说明）
├── dist/MD5校验工具.exe   # 打包产物（单文件，免环境直接运行）
├── 启动MD5校验工具.bat     # 启动脚本（优先 exe，兼容源码运行）
└── 使用说明.md            # 中文使用手册
```

## 说明

- 同一时刻只运行一个实例：重复双击会自动把已打开的窗口调到前台
- 同一次运行内重复拖入的文件会自动跳过
- 读取失败的文件（被独占占用等）会单独标红，不影响其他文件
- 缺少 WebView2 / 依赖时程序会弹窗提示，并在目录下生成「启动错误.log」
- 测试辅助参数：`--demo 文件路径…`（启动即计算）、`--autoquit 秒数`（自动退出）
