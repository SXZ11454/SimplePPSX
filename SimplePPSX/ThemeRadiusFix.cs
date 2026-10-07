using System;
using System.Collections.Generic;
using System.Windows;
using ModernWpf;

namespace SimplePPSX
{
	/// <summary>
	/// Square corners for the theme radii ModernWpf hands out to elements itself.
	///
	/// Style_\SquareCorners.xaml already replaces every radius key at application level, but a
	/// few templates carry their own resource dictionary with theme dictionaries and resolve
	/// those keys *inside* the element - the NumberBox spin popup is one (its root Grid
	/// declares ThemeManager.HasThemeResources="True"). Those element dictionaries are filled
	/// from the ModernWpf theme resources and keep the stock 4 / 2 radii.
	///
	/// Two hooks close that gap:
	/// - every framework element that comes alive with such a dictionary is patched right
	///   away (the values are looked up through the dictionary at run time, so a later
	///   change is picked up by whatever is already on screen),
	/// - the known elements are patched again whenever the application theme changes,
	///   because the switch copies fresh radii into the dictionaries that already exist.
	/// </summary>
	internal static class ThemeRadiusFix
	{
		/// <summary>
		/// The keys Style_\SquareCorners.xaml defines as 0. Must stay in step with
		/// that file: a key listed here but not there is harmless, and one listed
		/// there but not here would keep the stock radius on elements that carry
		/// their own dictionary, which is exactly the case this class exists for.
		///
		/// The two slider keys are deliberately absent, matching that file: sliders are
		/// not part of the square-corner rule. Pinning them here as well would keep
		/// re-imposing 0 on any element that carries its own theme dictionary, which
		/// is the one path that could not be undone by editing the resource alone.
		/// </summary>
		private static readonly string[] RadiusKeys =
		{
			"ControlCornerRadius",
			"OverlayCornerRadius",
			"PopupCornerRadius",
			"GridViewItemCornerRadius"
		};

		private static readonly CornerRadius Zero = new CornerRadius(0, 0, 0, 0);

		/// <summary>Elements carrying a patched dictionary (weak: the settings window is
		/// rebuilt on every open and the old instances must not be kept alive).</summary>
		private static readonly List<WeakReference> Patched = new List<WeakReference>();

		/// <summary>
		/// Kept in a static field so the delegate handed to ModernWpf never gets
		/// collected while the element it is attached to is still alive.
		/// </summary>
		private static RoutedEventHandler _themeChangedHandler;

		private static bool _hooked;

		/// <summary>Installs the element hook. Called once, before the first window is shown.</summary>
		internal static void Hook()
		{
			if (_hooked)
			{
				return;
			}

			_hooked = true;
			EventManager.RegisterClassHandler(
				typeof(FrameworkElement),
				FrameworkElement.LoadedEvent,
				new RoutedEventHandler(OnElementLoaded));
			Apply();
		}

		/// <summary>
		/// Re-applies the square radii whenever ModernWpf switches the theme, because the
		/// switch refills the dictionaries of the elements that are already on screen.
		///
		/// ActualThemeChanged is a routed event, so it has to be attached to a real
		/// element: Windows cannot be watched for theme changes directly, the application
		/// object is not an element. MainWindow and SettingsWindow watch themselves,
		/// which between them live for the whole session.
		/// </summary>
		internal static void Watch(FrameworkElement element)
		{
			if (element == null)
			{
				return;
			}
			if (_themeChangedHandler == null)
			{
				_themeChangedHandler = OnActualThemeChanged;
			}
			ThemeManager.AddActualThemeChangedHandler(element, _themeChangedHandler);
		}

		private static void OnActualThemeChanged(object sender, RoutedEventArgs e)
		{
			Apply();
			// A theme switch rebuilds parts of the visual tree, so anything that was
			// constructed before it has to be faded onto the new chrome again.
			ToolbarAppearance.Apply();
		}

		/// <summary>
		/// Patches the application resources and every element that was patched before -
		/// run after a theme switch, which refills those dictionaries with fresh values.
		/// </summary>
		internal static void Apply()
		{
			if (Application.Current == null)
			{
				return;
			}

			Patch(Application.Current.Resources);

			for (int i = Patched.Count - 1; i >= 0; i--)
			{
				FrameworkElement element = Patched[i].Target as FrameworkElement;
				if (element == null)
				{
					Patched.RemoveAt(i); // collected: drop the reference
				}
				else if (element.Resources != null)
				{
					Patch(element.Resources);
				}
			}
		}

		private static void OnElementLoaded(object sender, RoutedEventArgs e)
		{
			// Only the elements ModernWpf hands theme resources to can carry the stock
			// radii: they are the ones flagged with ThemeManager.HasThemeResources (the
			// root Grid of the number box spin popup is one) plus the popup roots. Every
			// other element resolves the keys from the application resources, which are
			// square already - and reading Resources on it would build an empty dictionary
			// that was never there.
			FrameworkElement element = e.OriginalSource as FrameworkElement;
			if (element == null)
			{
				return;
			}

			if (!ThemeManager.GetHasThemeResources(element) && !IsPopupRoot(element))
			{
				return;
			}

			// Remember the element when it carries radius keys at all - even if they are
			// square right now, because a theme switch may refill them later.
			if (element.Resources != null && Patch(element.Resources))
			{
				if (Patched.Count >= 256)
				{
					// rare (only ModernWpf templates declaring their own theme resources):
					// forget the entries that are gone before adding another one
					for (int i = Patched.Count - 1; i >= 0; i--)
					{
						if (Patched[i].Target == null)
						{
							Patched.RemoveAt(i);
						}
					}
				}

				Patched.Add(new WeakReference(element));
			}
		}

		/// <summary>The popup root (it may carry the theme resources of the popup as well;
		/// the WPF class itself is not public, so it is recognized by name).</summary>
		private static bool IsPopupRoot(DependencyObject element)
		{
			return element.GetType().Name == "PopupRoot";
		}

		/// <summary>
		/// Sets every non-zero corner radius of one dictionary (and its parts) to 0.
		/// Returns true when the dictionary holds radius keys, whatever their value is.
		/// </summary>
		private static bool Patch(ResourceDictionary dictionary)
		{
			bool hasRadiusKeys = false;

			for (int i = 0; i < RadiusKeys.Length; i++)
			{
				string key = RadiusKeys[i];
				if (!dictionary.Contains(key))
				{
					continue;
				}
				hasRadiusKeys = true;
				object value = dictionary[key];
				if (value is CornerRadius && (CornerRadius)value != Zero)
				{
					dictionary[key] = Zero;
				}
			}

			foreach (ResourceDictionary merged in dictionary.MergedDictionaries)
			{
				hasRadiusKeys |= Patch(merged);
			}

			// The theme dictionaries of ModernWpf's ResourceDictionaryEx are not part of
			// MergedDictionaries - they are the Light / Dark / HighContrast variants that
			// ThemeManager swaps while the app runs.
			ResourceDictionaryEx themed = dictionary as ResourceDictionaryEx;
			if (themed != null)
			{
				foreach (KeyValuePair<object, ResourceDictionary> pair in themed.ThemeDictionaries)
				{
					hasRadiusKeys |= Patch(pair.Value);
				}
			}

			return hasRadiusKeys;
		}
	}
}
