using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace SimplePPSX
{
	/// <summary>
	/// The system tray icon. The menu has exactly two entries: Settings and Exit.
	/// Everything else (start with Windows, preload, ...) moved into the settings
	/// window (SettingsWindow).
	///
	/// Depends on the already referenced System.Windows.Forms / System.Drawing, which
	/// are part of the Windows Desktop shared framework on every target framework.
	/// </summary>
	internal sealed class TrayIconService : IDisposable
	{
		// Where the startup entry lives (current user, no administrator needed)
		private const string RunKey = @"Software\Microsoft\Windows\CurrentVersion\Run";

		// Logical name of the icon embedded by the project file, see LoadIcon.
		private const string IconResourceName = "SPX.ico";

		private readonly NotifyIcon _notifyIcon;
		private bool _disposed;

		internal TrayIconService()
		{
			ToolStripMenuItem settingsMenuItem = new ToolStripMenuItem("设置");
			settingsMenuItem.Click += delegate { TrayIconService.ShowSettings(); };

			ToolStripMenuItem exitMenuItem = new ToolStripMenuItem("退出");
			exitMenuItem.Click += delegate { TrayIconService.RequestExit(); };

			ContextMenuStrip menu = new ContextMenuStrip();
			menu.Items.Add(settingsMenuItem);
			menu.Items.Add(exitMenuItem);

			_notifyIcon = new NotifyIcon();
			_notifyIcon.Icon = TrayIconService.LoadIcon();
			_notifyIcon.Text = "SimplePPSX";
			_notifyIcon.ContextMenuStrip = menu;
			_notifyIcon.Visible = true;
		}

		/// <summary>
		/// Opens the settings window. The tray menu click does not arrive on the UI
		/// thread, so it is marshalled over first (and the existing instance is
		/// reused, never a second window).
		/// </summary>
		private static void ShowSettings()
		{
			try
			{
				System.Windows.Application current = System.Windows.Application.Current;
				if (current == null)
				{
					return;
				}
				current.Dispatcher.BeginInvoke(new Action(delegate
				{
					SettingsWindow.ShowSingleton();
				}));
			}
			catch (Exception)
			{
				// the dispatcher is gone (shutting down): nothing to do
			}
		}

		/// <summary>Optional one-off balloon tip.</summary>
		internal void ShowBalloonTip(string title, string text)
		{
			try
			{
				_notifyIcon.BalloonTipTitle = title;
				_notifyIcon.BalloonTipText = text;
				_notifyIcon.ShowBalloonTip(3000);
			}
			catch (Exception)
			{
			}
		}

		/// <summary>
		/// Exit, from the tray menu.
		///
		/// The menu click does not arrive on the WPF UI thread, so it is marshalled
		/// over before anything is touched. (Touching WPF elements straight from the
		/// WinForms message loop was the other half of the exit being unreliable: the
		/// click handler ran, and the shutdown it asked for happened on the wrong
		/// thread or not at all.)
		///
		/// There is deliberately no early "nothing is open, just leave" branch:
		/// BottomArea.ExitAndShutdown closes the app either way, which is what this
		/// menu item promises. The on-screen exit button and the Escape key use
		/// ExitApplication instead and only end the presentation.
		/// </summary>
		private static void RequestExit()
		{
			try
			{
				System.Windows.Application app = System.Windows.Application.Current;
				if (app == null)
				{
					Environment.Exit(0);
					return;
				}
				app.Dispatcher.BeginInvoke(new Action(delegate
				{
					try
					{
						BottomArea bottomArea = ViewHelper.BottomArea;
						if (bottomArea != null)
						{
							bottomArea.ExitAndShutdown();
						}
						else
						{
							// No toolbar was ever built (nothing presented yet, preload off).
							// This file uses System.Windows.Forms, whose Application shadows
							// the WPF one of the same name, hence the full namespace.
							System.Windows.Application.Current.Shutdown();
						}
					}
					catch (Exception)
					{
						// Last resort: the user asked to quit, so quit.
						Environment.Exit(0);
					}
				}));
			}
			catch (Exception)
			{
				// the dispatcher is gone: already shutting down
			}
		}

		/// <summary>
		/// The application icon: the embedded copy first, then SPX.ico next to the
		/// executable. The embedded copy is what makes the single-file Win7 build
		/// work, because that publish produces the exe and nothing else.
		/// </summary>
		private static Icon LoadIcon()
		{
			try
			{
				using (Stream embedded = typeof(TrayIconService).Assembly.GetManifestResourceStream(TrayIconService.IconResourceName))
				{
					if (embedded != null)
					{
						return new Icon(embedded);
					}
				}
			}
			catch (Exception)
			{
			}
			try
			{
				string exeDir = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
				string ico = Path.Combine(exeDir, "SPX.ico");
				if (File.Exists(ico))
				{
					return new Icon(ico);
				}
			}
			catch (Exception)
			{
			}
			return SystemIcons.Application;
		}

		// ---- start with Windows (the settings switch uses the same read/write) ----

		private static string AutoStartValueName
		{
			get { return "SimplePPSX"; }
		}

		/// <summary>Path of the running executable.</summary>
		private static string GetExecutablePath()
		{
			try
			{
				Process main = Process.GetCurrentProcess();
				if (main.MainModule != null && !string.IsNullOrEmpty(main.MainModule.FileName))
				{
					return main.MainModule.FileName;
				}
			}
			catch (Exception)
			{
			}
			return Assembly.GetEntryAssembly().Location;
		}

		internal static bool IsAutoStartEnabled()
		{
			try
			{
				using (Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(RunKey))
				{
					if (key == null)
					{
						return false;
					}
					return key.GetValue(TrayIconService.AutoStartValueName) != null;
				}
			}
			catch (Exception)
			{
				return false;
			}
		}

		internal static bool SetAutoStart(bool enabled)
		{
			try
			{
				using (Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(RunKey))
				{
					if (key == null)
					{
						return false;
					}
					if (enabled)
					{
						string path = TrayIconService.GetExecutablePath();
						if (string.IsNullOrEmpty(path))
						{
							return false;
						}
						// A path containing spaces has to be quoted
						key.SetValue(TrayIconService.AutoStartValueName, "\"" + path + "\"");
					}
					else
					{
						key.DeleteValue(TrayIconService.AutoStartValueName, false);
					}
				}
				return true;
			}
			catch (Exception)
			{
				return false;
			}
		}

		public void Dispose()
		{
			if (this._disposed)
			{
				return;
			}
			this._disposed = true;
			try
			{
				_notifyIcon.Visible = false;
				if (_notifyIcon.ContextMenuStrip != null)
				{
					_notifyIcon.ContextMenuStrip.Dispose();
					_notifyIcon.ContextMenuStrip = null;
				}
				_notifyIcon.Dispose();
			}
			catch (Exception)
			{
			}
		}
	}
}
