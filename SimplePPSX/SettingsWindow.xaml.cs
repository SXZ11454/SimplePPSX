using System;
using System.Windows;

namespace SimplePPSX
{
	/// <summary>
	/// The settings window. This is what the tray menu's "Settings" entry opens, and
	/// only one instance ever exists.
	///
	/// The presentation simulation window that used to sit behind this one (the live
	/// preview of the toolbar positions) was removed together with the position
	/// settings: the toolbars are always pinned to the bottom edge now, so there was
	/// nothing to preview.
	/// </summary>
	public partial class SettingsWindow : Window
	{
		private static SettingsWindow _instance;

		public SettingsWindow()
		{
			this.InitializeComponent();
			this.Page.Attach(this);

			this.SourceInitialized += delegate
			{
				// Windows 11 rounds window corners on its own; ask it not to, so this
				// window matches the rest of the square-corner design
				// (Style_\SquareCorners.xaml).
				WindowCorners.DisableRounding(this);
			};
			// Keep the square radii in step with theme changes for as long as this
			// window is open.
			ThemeRadiusFix.Watch(this);
			this.Closed += delegate
			{
				if (SettingsWindow._instance == this)
				{
					SettingsWindow._instance = null;
				}
			};
		}

		/// <summary>Tray entry point: activate the existing instance or create one.</summary>
		internal static void ShowSingleton()
		{
			if (SettingsWindow._instance != null)
			{
				if (SettingsWindow._instance.WindowState == WindowState.Minimized)
				{
					SettingsWindow._instance.WindowState = WindowState.Normal;
				}
				SettingsWindow._instance.Activate();
				return;
			}

			SettingsWindow window = new SettingsWindow();
			SettingsWindow._instance = window;
			window.Show();
		}
	}
}