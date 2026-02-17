# CaretClip

Win11 风格剪贴板小工具（WPF）：
- 记录剪贴板历史（文本、图片、文件列表）
- 查看并管理历史内容
- 编辑文本历史并保存
- 设置保存时长（1 天 / 7 天 / 30 天 / 永久）
- 全局快捷键 `Alt+Shift+V` 唤出或隐藏小窗口（窗口置顶）
- 系统托盘常驻（双击托盘图标可唤出）
- 开机启动开关（界面和托盘菜单都可控制）
- 历史搜索与条目置顶（PINNED）
- 主题切换（跟随系统 / 浅色 / 深色）
- Win11 窗口效果（圆角、Mica 背景、显示隐藏动画）
- 主窗口仅保留复制相关内容，设置与关于集中在“设置”面板

## 运行

```powershell
dotnet build
dotnet run
```

## 数据存储

数据保存在：

`%LOCALAPPDATA%\CaretClip`

- `history.json`
- `settings.json`
