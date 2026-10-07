# SimplePPSX

**English** | [简体中文](README.zh-CN.md)

A floating toolbar overlay for Microsoft PowerPoint slide shows on Windows.

SimplePPSX sits on top of a running PowerPoint presentation as a transparent borderless window and gives you pen tools, thumbnail navigation and show controls without covering your slides.

## Features

- **Overlay that follows the show** — the toolbars float in when a slide show starts and sink away when it ends; the app itself lives in the tray.
- **Five toolbar strips** — left, left-bottom, center, right and right-bottom, each collapsible together with the thumbnail panels so you can get a completely clean screen.
- **Tools**
  - *Pointer* — hand the cursor back to PowerPoint.
  - *Pen* — draw annotations; tap the button again to open the ink color panel.
  - *Eraser* — a slider sizes the eraser; push the slider to the end to wipe every ink stroke on the current slide (the panel closes first, then the ink is cleared).
- **Thumbnail panels** on both sides of the screen — scroll through the deck and click a thumbnail to jump to that slide.
- **Slide navigation** — previous/next slide from the keyboard or the toolbar, slide counter included.
- **Tray-resident and single-instance** — the tray menu's Exit ends a running show first, then quits.
- **Settings window** — start with Windows, preload the overlay window, and tune the overlay opacity.
- **Logs** — written next to the executable under `Program\SimplePPSX\Log\` (`debug.txt` / `error.txt`).

## Requirements

- Windows 7 SP1 or later for the .NET 6 build, Windows 10 (1607) or later for the .NET 10 build.
- Microsoft PowerPoint (Windows).
- Self-contained builds need no pre-installed runtime; the framework-dependent build needs the .NET Desktop Runtime.

## Build from source

Open `SimplePPSX.sln` in Visual Studio (with the .NET desktop development workload) and build, or use the CLI:

```powershell
dotnet build SimplePPSX.sln -c Debug
```

The project multi-targets `net10.0-windows` and `net6.0-windows` and always builds for `win-x86`.

## Publish

Each profile writes its output to `SimplePPSX\bin\Release\<framework>\win-x86\publish\`:

```powershell
# .NET 6, self-contained single file (Windows 7 compatible)
dotnet publish SimplePPSX\SimplePPSX.csproj -c Release -p:TargetFramework=net6.0-windows -p:PublishProfile=Win7-Net6-x86

# .NET 10, self-contained single file
dotnet publish SimplePPSX\SimplePPSX.csproj -c Release -p:TargetFramework=net10.0-windows -p:PublishProfile=Net10-SelfContained-x86

# .NET 10, framework-dependent (small, needs the .NET Desktop Runtime)
dotnet publish SimplePPSX\SimplePPSX.csproj -c Release -p:TargetFramework=net10.0-windows -p:PublishProfile=Net10-x86
```

## Usage

1. Start SimplePPSX — it stays in the notification area.
2. Start (or switch to) a PowerPoint slide show — the five strips float up over the slides.
3. Drive the show from the strips at the bottom of the screen:
   - **Pointer / Pen / Eraser** switch tools; the color panel and the eraser size slider open above the strip.
   - **Collapse / expand** folds the strips (and the thumbnails) away.
   - **End show** stops the presentation; everything sinks out together.
4. Use the left/right thumbnail panel to browse and jump between slides.
5. `Esc` ends the slide show but keeps the app in the tray.

### Shortcuts

| Key | Action |
| --- | ------ |
| `Page Up` / `←` / `↑` | Previous slide |
| `Page Down` / `→` / `↓` | Next slide |
| `Esc` | End the slide show |

## Settings

| Setting | Effect |
| ------- | ------ |
| Start automatically | Launch SimplePPSX at sign-in |
| Preload | Create the overlay window up front so the first show appears instantly |
| Overlay opacity | Slider or exact value for the overlay window |

## Project layout

```text
SimplePPSX.sln                 Solution
SPX.ico                        Application icon
SimplePPSX\
  App.xaml(.cs)                Single-instance entry point, tray bootstrap
  MainWindow.xaml(.cs)         Full-screen overlay; show/hide choreography for every host
  BottomArea / LeftArea / LeftBottomArea / RightArea / RightBottomArea
                               The five toolbar strips (tools, navigation, collapse)
  SlideThumbnailPanel.xaml     Thumbnails used by the left/right panels
  PptOperator.cs               PowerPoint COM automation (win-x86)
  TrayIconService.cs           Tray icon and menu
  ToolbarAnimation.cs          Enter/exit motion built on pure translate animations
  SettingsWindow / SettingsPage
  Style_\                      Control styles and templates
  Properties\PublishProfiles\  The publish profiles used above
  Microsoft\, LinqToVisualTree\  Interop and tree-walking helpers
```

## License

Licensed under the [Apache License 2.0](LICENSE).
