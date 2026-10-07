using System;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using ModernWpf.Controls;

namespace SimplePPSX
{
	/// <summary>
	/// The settings page, hosted by <see cref="SettingsWindow"/>.
	///
	/// Every control applies and saves its change immediately, so there is no
	/// OK / Cancel:
	///   - start with Windows, preload window: registry / static switch,
	///   - toolbar opacity: changes <see cref="ToolbarAppearance.Opacity"/>, which
	///     reaches every backdrop.
	///
	/// The toolbar positions and the presentation simulation window were removed:
	/// the three toolbars are always pinned to the bottom edge of the presentation
	/// area now, so there was nothing left to configure.
	/// </summary>
	public partial class SettingsPage : UserControl
	{
		/// <summary>The owning window, used for the close button and the hint dialog.</summary>
		private SettingsWindow _owner;

		/// <summary>Set while the controls are being filled, so restoring the saved
		/// values is not mistaken for the user changing something.</summary>
		private bool _loading = true;

		public SettingsPage()
		{
			this.InitializeComponent();
			this.DataContext = this;

			this.LoadValues();

			this._loading = false;
		}

		/// <summary>Called by the window once it exists, so the page can reach it.</summary>
		internal void Attach(SettingsWindow owner)
		{
			this._owner = owner;
		}

		// ================= load =================

		private void LoadValues()
		{
			this.AutoStartSwitch.IsOn = TrayIconService.IsAutoStartEnabled();
			this.PreloadSwitch.IsOn = ViewHelper.IsPreloadEnabled;

			double percent = ToolbarAppearance.OpacityPercent;
			this.OpacitySlider.Value = percent;
			this.OpacityNumberBox.Value = percent;
		}

		// ================= start with Windows =================

		private void AutoStartSwitch_Toggled(object sender, RoutedEventArgs e)
		{
			if (this._loading)
			{
				return;
			}
			bool target = this.AutoStartSwitch.IsOn;
			if (TrayIconService.SetAutoStart(target))
			{
				return;
			}
			// Writing the registry failed: flip the switch back and say why.
			this._loading = true;
			this.AutoStartSwitch.IsOn = !target;
			this._loading = false;
			this.ShowHint("Could not update the startup entry. Please check your registry permissions.");
		}

		// ================= preload =================

		private void PreloadSwitch_Toggled(object sender, RoutedEventArgs e)
		{
			if (this._loading)
			{
				return;
			}
			ViewHelper.IsPreloadEnabled = this.PreloadSwitch.IsOn;
		}

		// ================= opacity =================

		// .NET Framework's Slider.ValueChanged uses RoutedPropertyChangedEventArgs<double>
		// (RangeBaseValueChangedEventArgs<T> only exists from .NET 5 on).
		private void OpacitySlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
		{
			if (this._loading || e.NewValue < 0)
			{
				return;
			}
			this._loading = true;
			this.OpacityNumberBox.Value = e.NewValue;
			this._loading = false;
			ToolbarAppearance.SetOpacityPercent(e.NewValue);
		}

		private void OpacityNumberBox_ValueChanged(object sender, NumberBoxValueChangedEventArgs e)
		{
			if (this._loading || e.NewValue < 0)
			{
				return;
			}
			this._loading = true;
			this.OpacitySlider.Value = e.NewValue;
			this._loading = false;
			ToolbarAppearance.SetOpacityPercent(e.NewValue);
		}

		private void ResetOpacityButton_Click(object sender, RoutedEventArgs e)
		{
			this._loading = true;
			this.OpacitySlider.Value = ToolbarAppearance.DefaultOpacity * 100.0;
			this.OpacityNumberBox.Value = this.OpacitySlider.Value;
			this._loading = false;
			ToolbarAppearance.Opacity = ToolbarAppearance.DefaultOpacity;
		}

		// ================= misc =================

		private void CloseButton_Click(object sender, RoutedEventArgs e)
		{
			if (this._owner != null)
			{
				this._owner.Close();
			}
		}

		/// <summary>One-off notice, shown with ModernWpfUI's ContentDialog.</summary>
		private void ShowHint(string message)
		{
			ContentDialog dialog = new ContentDialog();
			if (this._owner != null)
			{
				dialog.Owner = this._owner;
			}
			dialog.Title = "提示";
			dialog.Content = message;
			dialog.CloseButtonText = "知道了";
			try
			{
				dialog.ShowAsync();
			}
			catch (Exception)
			{
				// No usable owner window: stay quiet, the switch has already been put back.
			}
		}

		// ================= texts for the About section =================

		public string AppVersion
		{
			get
			{
				Assembly assembly = Assembly.GetExecutingAssembly();
				AssemblyInformationalVersionAttribute info = assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>();
				string version = info != null && !string.IsNullOrEmpty(info.InformationalVersion)
					? info.InformationalVersion
					: assembly.GetName().Version.ToString();
				AssemblyTitleAttribute title = assembly.GetCustomAttribute<AssemblyTitleAttribute>();
				string name = title != null && !string.IsNullOrEmpty(title.Title) ? title.Title : "SimplePPSX";
				return name + "  " + version;
			}
		}

		public string AppCopyright
		{
			get
			{
				Assembly assembly = Assembly.GetExecutingAssembly();
				AssemblyCopyrightAttribute copyright = assembly.GetCustomAttribute<AssemblyCopyrightAttribute>();
				return copyright != null ? copyright.Copyright : string.Empty;
			}
		}

		public string AppDescription
		{
			get
			{
				Assembly assembly = Assembly.GetExecutingAssembly();
				AssemblyDescriptionAttribute description = assembly.GetCustomAttribute<AssemblyDescriptionAttribute>();
				return description != null ? description.Description : "PowerPoint presentation toolbar";
			}
		}

		}
}
