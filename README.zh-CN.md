# SimplePPSX

[English](README.md) | **简体中文**

Windows 下 Microsoft PowerPoint 放映用的浮动工具栏覆盖层。

SimplePPSX 以透明无边框窗口悬浮在正在放映的 PowerPoint 之上，在不遮挡幻灯片的前提下提供画笔工具、缩略图导航和放映控制。

## 功能特性

- **跟随放映的覆盖层** — 放映开始时工具栏整排上浮进入，停止放映时一起下沉退场；应用本身常驻托盘。
- **五条工具栏** — 左、左下、中间、右、右下，可与缩略图面板一同折叠，随时获得完全干净的屏幕。
- **工具**
  - *指针* — 把光标交还给 PowerPoint。
  - *画笔* — 标注书写；再点一次按钮打开墨迹颜色面板（全部颜色分两行排布，滑动描边框指示当前色），按钮底部常显当前墨迹颜色。
  - *橡皮* — 滑块调节粗细；把滑块拖到头即清除当前页全部墨迹（先关闭面板，再清除墨迹）。
- **屏幕两侧缩略图面板** — 滚动浏览整个演示文稿，点击缩略图直接跳页；页码数字排在缩略图旁边。
- **深浅色模式** — 界面默认跟随 Windows，也可在设置中固定浅色或深色，工具栏配色一起切换。
- **翻页控制** — 键盘或工具栏翻页，附带页码计数。
- **托盘常驻、单实例** — 托盘菜单的"退出"会先结束正在进行的放映，再退出程序。
- **设置窗口** — 开机自启、预加载覆盖层窗口、预获取缩略图、选择深浅模式、调节覆盖层不透明度。
- **日志** — 写在可执行文件旁的 `Program\SimplePPSX\Log\`（`debug.txt` / `error.txt`）。

## 运行要求

- .NET 6 版本：Windows 7 SP1 及以上；.NET 10 版本：Windows 10（1607）及以上。
- Microsoft PowerPoint（Windows）。
- 自包含版本无需预装运行时；框架依赖版本需要 .NET 桌面运行时。

## 从源码构建

用 Visual Studio（含 .NET 桌面开发工作负载）打开 `SimplePPSX.sln` 生成，或使用命令行：

```powershell
dotnet build SimplePPSX.sln -c Debug
```

项目同时面向 `net10.0-windows` 与 `net6.0-windows`，始终以 `win-x86` 为目标。

## 发布

每个发布配置的产物输出到 `SimplePPSX\bin\Release\<framework>\win-x86\publish\`：

```powershell
# .NET 6 自包含单文件（兼容 Windows 7）
dotnet publish SimplePPSX\SimplePPSX.csproj -c Release -p:TargetFramework=net6.0-windows -p:PublishProfile=Win7-Net6-x86

# .NET 10 自包含单文件
dotnet publish SimplePPSX\SimplePPSX.csproj -c Release -p:TargetFramework=net10.0-windows -p:PublishProfile=Net10-SelfContained-x86

# .NET 10 框架依赖（体积小，需 .NET 桌面运行时）
dotnet publish SimplePPSX\SimplePPSX.csproj -c Release -p:TargetFramework=net10.0-windows -p:PublishProfile=Net10-x86
```

## 使用方法

1. 启动 SimplePPSX —— 程序驻留在通知区域。
2. 启动（或切换到）PowerPoint 放映 —— 五条工具栏自动上浮到幻灯片上方。
3. 用屏幕底部的工具栏控制放映：
   - **指针 / 画笔 / 橡皮** 切换工具；颜色面板与橡皮粗细滑块会在工具栏上方打开。
   - **收起 / 展开** 把工具栏（连同缩略图）折叠起来。
   - **退出放映** 结束当前演示，所有条目一起下沉退场。
4. 用左/右缩略图面板浏览并跳转幻灯片。
5. `Esc` 结束放映，程序保留在托盘。

### 快捷键

| 按键 | 作用 |
| --- | ---- |
| `Page Up` / `←` / `↑` | 上一页 |
| `Page Down` / `→` / `↓` | 下一页 |
| `Esc` | 结束放映 |

## 设置项

| 设置 | 作用 |
| --- | ---- |
| 开机自启 | 登录 Windows 时自动启动 |
| 预加载 | 提前创建覆盖层窗口，让第一场放映立即显示 |
| 预获取缩略图 | 放映开始即在后台导出全部缩略图，面板首次打开即完整；关闭则推迟到首次打开时导出 |
| 深浅模式 | 跟随 Windows，或固定浅色 / 深色配色 |
| 覆盖层不透明度 | 滑块或精确数值调节覆盖层窗口透明度 |

## 项目结构

```text
SimplePPSX.sln                 解决方案
SPX.ico                        应用图标
SimplePPSX\
  App.xaml(.cs)                单实例入口、托盘引导
  MainWindow.xaml(.cs)         全屏覆盖层；各宿主窗口的进出场调度
  BottomArea / LeftArea / LeftBottomArea / RightArea / RightBottomArea
                               五条工具栏（工具、导航、折叠）
  SlideThumbnailPanel.xaml     左右面板使用的缩略图控件
  PptOperator.cs               PowerPoint COM 自动化（win-x86）
  TrayIconService.cs           托盘图标与菜单
  ToolbarAnimation.cs          基于纯位移动画的进出场
  SettingsWindow / SettingsPage
  Style_\                      控件样式与模板
  Properties\PublishProfiles\  上文使用的发布配置
  Microsoft\, LinqToVisualTree\  互操作与树遍历辅助
```

## 许可证

基于 [Apache License 2.0](LICENSE) 开源。内置的 Material Symbols Sharp 图标字体同样采用 Apache License 2.0 许可。
