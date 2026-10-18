using System;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using ModernWpf;
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
			this.SelectDecodeWidth(ViewHelper.ThumbnailDecodeWidth);
			this.SelectTheme(ViewHelper.ThemeMode);

			double percent = ToolbarAppearance.OpacityPercent;
			double idlePercent = ToolbarAppearance.IdleOpacityPercent;
			this.IdleOpacitySlider.Value = idlePercent;
			this.IdleOpacityNumberBox.Value = idlePercent;
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

		// ================= thumbnail decode width =================

		/// <summary>Highlights the item carrying the stored pixel width; unknown falls back to the default.</summary>
		private void SelectDecodeWidth(int width)
		{
			ComboBoxItem match = this.FindDecodeWidthItem(width);
			if (match == null)
			{
				match = this.FindDecodeWidthItem(ConfigController.DefaultThumbnailDecodeWidth);
			}
			if (match != null)
			{
				this.DecodeWidthComboBox.SelectedItem = match;
			}
			else
			{
				this.DecodeWidthComboBox.SelectedIndex = 1;
			}
		}

		private ComboBoxItem FindDecodeWidthItem(int width)
		{
			foreach (object item in this.DecodeWidthComboBox.Items)
			{
				ComboBoxItem box = item as ComboBoxItem;
				int tag;
				if (box != null && box.Tag is string && int.TryParse((string)box.Tag, out tag) && tag == width)
				{
					return box;
				}
			}
			return null;
		}

		private void DecodeWidthComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
		{
			if (this._loading)
			{
				return;
			}
			ComboBoxItem selected = this.DecodeWidthComboBox.SelectedItem as ComboBoxItem;
			int width;
			if (selected == null || !(selected.Tag is string) || !int.TryParse((string)selected.Tag, out width))
			{
				width = ConfigController.DefaultThumbnailDecodeWidth;
			}
			// Saved like every other control: the next LoadThumbnail picks it up,
			// so the new width applies from the next presentation with no restart.
			ViewHelper.ThumbnailDecodeWidth = width;
			Log.WriteDebug("settings: thumbnail decode width set to " + width + (width <= 0 ? " (native)" : "px"));
		}

		// ================= theme =================

		/// <summary>Highlights the item matching the stored mode; unknown falls back to system.</summary>
		private void SelectTheme(string mode)
		{
			foreach (object item in this.ThemeComboBox.Items)
			{
				ComboBoxItem box = item as ComboBoxItem;
				if (box != null && string.Equals(box.Tag as string, mode, StringComparison.OrdinalIgnoreCase))
				{
					this.ThemeComboBox.SelectedItem = box;
					return;
				}
			}
			this.ThemeComboBox.SelectedIndex = 0;
		}

		private void ThemeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
		{
			if (this._loading)
			{
				return;
			}
			ComboBoxItem selected = this.ThemeComboBox.SelectedItem as ComboBoxItem;
			string mode = selected != null && selected.Tag is string ? (string)selected.Tag : ToolbarTheme.ModeSystem;
			ViewHelper.ThemeMode = mode;
			ThemeManager.Current.ApplicationTheme = ToolbarTheme.ToApplicationTheme(mode);
			// The brush set lives in the application resources and is stamped
			// from code, so restamp it here too - the windows' actual-theme
			// handler does it as well, and the second write is the same values.
			ToolbarTheme.Apply();
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
			int storedPercent = (int)Math.Round(e.NewValue);
			ViewHelper.ToolbarOpacityPercent = storedPercent;
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
			int storedPercent = (int)Math.Round(e.NewValue);
			ViewHelper.ToolbarOpacityPercent = storedPercent;
			ToolbarAppearance.SetOpacityPercent(e.NewValue);
		}

		private void IdleOpacitySlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
		{
			if (this._loading || e.NewValue < 0)
			{
				return;
			}
			this._loading = true;
			this.IdleOpacityNumberBox.Value = e.NewValue;
			this._loading = false;
			int storedPercent = (int)Math.Round(e.NewValue);
			ViewHelper.ToolbarIdleOpacityPercent = storedPercent;
			ToolbarAppearance.SetIdleOpacityPercent(e.NewValue);
		}

		private void IdleOpacityNumberBox_ValueChanged(object sender, NumberBoxValueChangedEventArgs e)
		{
			if (this._loading || e.NewValue < 0)
			{
				return;
			}
			this._loading = true;
			this.IdleOpacitySlider.Value = e.NewValue;
			this._loading = false;
			int storedPercent = (int)Math.Round(e.NewValue);
			ViewHelper.ToolbarIdleOpacityPercent = storedPercent;
			ToolbarAppearance.SetIdleOpacityPercent(e.NewValue);
		}

		private void ResetOpacityButton_Click(object sender, RoutedEventArgs e)
		{
			this._loading = true;
			this.OpacitySlider.Value = ToolbarAppearance.DefaultOpacity * 100.0;
			this.OpacityNumberBox.Value = this.OpacitySlider.Value;
			this.IdleOpacitySlider.Value = ToolbarAppearance.DefaultIdleOpacity * 100.0;
			this.IdleOpacityNumberBox.Value = this.IdleOpacitySlider.Value;
			this._loading = false;
			int storedDefault = (int)Math.Round(ToolbarAppearance.DefaultOpacity * 100.0);
			int storedIdle = (int)Math.Round(ToolbarAppearance.DefaultIdleOpacity * 100.0);
			ViewHelper.ToolbarOpacityPercent = storedDefault;
			ToolbarAppearance.SetOpacityPercent(storedDefault);
			ViewHelper.ToolbarIdleOpacityPercent = storedIdle;
			ToolbarAppearance.SetIdleOpacityPercent(storedIdle);
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
