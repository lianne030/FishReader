# FishReader

FishReader 是一个适用于 Windows 的透明悬浮小说阅读器。解压后双击 `启动阅读器.cmd` 即可使用，无需安装开发环境。

## 功能

- 无边框透明窗口，可置顶、拖动和调整显示区域
- 鼠标滚轮、方向键与 PageUp/PageDown 翻页
- 自动识别常见中文章节标题，支持上一章和下一章
- 独立记录每本书的阅读进度，重启后自动续读
- 字号、字体颜色、行间距和段间距调节
- 鼠标穿透与全局快捷键
- 支持 TXT、EPUB、HTML、Markdown、FB2、RTF、DOCX 文本内容
- 所有数据只保存在本地

## 立即使用

1. 从仓库中的 [`dist/FishReader-portable.zip`](https://github.com/lianne030/FishReader/raw/refs/heads/main/dist/FishReader-portable.zip) 下载便携包。
2. 解压到任意文件夹。
3. 双击 `启动阅读器.cmd`。
4. 在透明窗口内右键，选择“打开小说”。

Windows SmartScreen 或安全软件首次可能提示未知脚本。项目源码完全包含在 `FishReader.ps1` 中，可直接查看。

## 快捷操作

- 左键拖动：移动窗口
- 双击窗口：保存进度并关闭
- 滚轮向下/向上：下一页/上一页
- `Ctrl + Alt + H`：保存进度并关闭
- `Ctrl + Alt + T`：切换鼠标穿透
- `Esc`：保存进度并关闭

详细说明见 [使用说明.md](使用说明.md)。

## 数据与隐私

阅读记录和设置存放在程序目录下的 `FishReaderData`。仓库及发布包不包含任何电子书、个人书单、阅读记录或日志。

## 系统要求

- Windows 10 或 Windows 11
- 系统自带 Windows PowerShell 5.1 与 WPF

## 许可证

[MIT License](LICENSE)
