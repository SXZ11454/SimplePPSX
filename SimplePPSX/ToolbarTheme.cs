using System;
using System.Windows;
using System.Windows.Media;
using Microsoft.Win32;
using ModernWpf;

namespace SimplePPSX
{
	/// <summary>
	/// The light / dark brush set for every piece of toolbar chrome, and the
	/// plumbing that chooses between them.
	///
	/// Style_\ToolbarStyle.xaml used to declare the ten Toolbar*Brush resources
	/// once, in the light colours, and they never changed after parse. They are
	/// stamped from here now: two parallel colour tables, and <see cref="Apply"/>
	/// writes the pair for the actual theme into the application resources under
	/// the same keys.
	///
	/// Why from code and not from XAML theme dictionaries: WPF dictionaries have
	/// no Light / Dark variants of their own. ModernWpf's ResourceDictionaryEx
	/// does, but only the library's own dictionaries are wired for it - a plain
	/// entry in this app's resources never changes. Every consumer already
	/// resolves the brushes with DynamicResource, so replacing the entries under
	/// the keys repaints whatever is on screen without a single element knowing.
	///
	/// When it runs: once at start-up (App.OnStartup, after the theme mode is
	/// pushed onto ModernWpf) and again on every actual theme change - the
	/// windows watch for that through ThemeRadiusFix, which calls
	/// <see cref="Apply"/> before re-fading the chrome.
	///
	/// The mode itself (system / light / dark) is the settings page's
	/// ThemeComboBox, stored through the usual chain - see ViewHelper.ThemeMode.
	/// </summary>
	internal static class ToolbarTheme
	{
		/// <summary>The three modes, as the settings page stores them.</summary>
		internal const string ModeSystem = "system";
		internal const string ModeLight = "light";
		internal const string ModeDark = "dark";

		private static readonly string[] Keys =
		{
			"ToolbarBackgroundBrush",
			"ToolbarBorderBrush",
			"ToolbarDividerBrush",
			"ToolbarForegroundBrush",
			"ToolbarFillSelectedBrush",
			"ToolbarForegroundSelectedBrush",
			"ToolbarAccentOutlineBrush",
			"ToolbarFillHoverBrush",
			"ToolbarFillPressedBrush",
			"ToolbarTrackPlateBrush"
		};

		// Light: the old design, kept value for value - white plate, #666 outline
		// and hairlines, #3F3F3F ink, the pale track for the eraser slider. The
		// accent pair is #0078D4, the hover and pressed washes are black-alpha.
		private static readonly Color[] Light =
		{
			Color.FromRgb(0xFF, 0xFF, 0xFF), // ToolbarBackgroundBrush
			Color.FromRgb(0x66, 0x66, 0x66), // ToolbarBorderBrush
			Color.FromRgb(0x66, 0x66, 0x66), // ToolbarDividerBrush
			Color.FromRgb(0x3F, 0x3F, 0x3F), // ToolbarForegroundBrush
			Color.FromRgb(0x00, 0x78, 0xD4), // ToolbarFillSelectedBrush
			Color.FromRgb(0xFF, 0xFF, 0xFF), // ToolbarForegroundSelectedBrush
			Color.FromRgb(0x00, 0x78, 0xD4), // ToolbarAccentOutlineBrush
			Color.FromArgb(0x1A, 0x00, 0x00, 0x00), // ToolbarFillHoverBrush
			Color.FromArgb(0x26, 0x00, 0x00, 0x00), // ToolbarFillPressedBrush
			Color.FromRgb(0xEC, 0xEC, 0xEC)  // ToolbarTrackPlateBrush
		};

		// Dark: a #2B plate under the same toolbar opacity, with the ink, the
		// outline and the washes turned around it - light ink on the dark plate,
		// white-alpha for hover and pressed, a #3B track. The accent pair stays
		// the light set's: #0078D4 reads on the dark plate exactly as it does on
		// white, and the selected glyph is white in both.
		private static readonly Color[] Dark =
		{
			Color.FromRgb(0x2B, 0x2B, 0x2B), // ToolbarBackgroundBrush
			Color.FromRgb(0x9B, 0x9B, 0x9B), // ToolbarBorderBrush
			Color.FromRgb(0x9B, 0x9B, 0x9B), // ToolbarDividerBrush
			Color.FromRgb(0xEC, 0xEC, 0xEC), // ToolbarForegroundBrush
			Color.FromRgb(0x00, 0x78, 0xD4), // ToolbarFillSelectedBrush
			Color.FromRgb(0xFF, 0xFF, 0xFF), // ToolbarForegroundSelectedBrush
			Color.FromRgb(0x00, 0x78, 0xD4), // ToolbarAccentOutlineBrush
			Color.FromArgb(0x1A, 0xFF, 0xFF, 0xFF), // ToolbarFillHoverBrush
			Color.FromArgb(0x26, 0xFF, 0xFF, 0xFF), // ToolbarFillPressedBrush
			Color.FromRgb(0x3B, 0x3B, 0x3B)  // ToolbarTrackPlateBrush
		};

		/// <summary>
		/// Writes the brush pair for the current theme into the application
		/// resources, under the same keys ToolbarStyle.xaml declared. Harmless to
		/// call twice - the second write is the same values.
		/// </summary>
		internal static void Apply()
		{
			if (Application.Current == null)
			{
				return;
			}

			Color[] palette = ToolbarTheme.IsDark() ? Dark : Light;
			ResourceDictionary resources = Application.Current.Resources;
			for (int i = 0; i < ToolbarTheme.Keys.Length; i++)
			{
				resources[ToolbarTheme.Keys[i]] = new SolidColorBrush(palette[i]);
			}
		}

		/// <summary>
		/// Maps the stored mode onto the ModernWpf application theme. The library
		/// takes a nullable: null is "no preference", which is what follows the
		/// system - the same as UWP's Default.
		/// </summary>
		internal static ApplicationTheme? ToApplicationTheme(string mode)
		{
			if (mode == ToolbarTheme.ModeLight)
			{
				return ApplicationTheme.Light;
			}
			if (mode == ToolbarTheme.ModeDark)
			{
				return ApplicationTheme.Dark;
			}
			return null;
		}

		/// <summary>
		/// The resolved colour side, from the stored mode. "system" asks Windows
		/// directly - the sameApps Use Light value ModernWpf itself follows -
		/// because ThemeManager offers no application-level ActualTheme to read,
		/// and re-reading it here keeps the brush set in step with a live system
		/// switch (the windows' ActualTheme change event calls Apply again).
		/// </summary>
		private static bool IsDark()
		{
			string mode = ViewHelper.ThemeMode;
			if (mode == ToolbarTheme.ModeDark)
			{
				return true;
			}
			if (mode == ToolbarTheme.ModeLight)
			{
				return false;
			}

			// system: whatever Windows has the app theme set to. The value is 1
			// for light and 0 for dark; a missing key is a Windows too old to
			// have the setting, which only ever was light.
			try
			{
				using (RegistryKey key = Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize"))
				{
					object value = key != null ? key.GetValue("AppsUseLightTheme") : null;
					if (value is int)
					{
						return (int)value == 0;
					}
				}
			}
			catch (Exception)
			{
				// A registry read is best effort; fall through to the default.
			}

			return false;
		}
	}
}
