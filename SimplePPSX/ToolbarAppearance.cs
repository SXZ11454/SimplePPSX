using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Media;

namespace SimplePPSX
{
	/// <summary>
	/// The single opacity control for the presentation toolbars.
	///
	/// This is the extension point the settings window drives. No toolbar writes a
	/// colour of its own: every piece of chrome (the backdrops, the 1px outlines and
	/// the hairlines between buttons) is a plain Rectangle filled from
	/// ToolbarBackgroundBrush / ToolbarBorderBrush / ToolbarDividerBrush, and the
	/// controls register those Rectangles here. Changing <see cref="Opacity"/> then
	/// fades every one of them at once, which means
	///   - the toolbars do not need to know about each other,
	///   - the settings window changes one value and every toolbar follows.
	///
	/// Why element Opacity and not brush Opacity: WPF freezes every freezable that
	/// comes out of a resource dictionary, and ModernWpf's ThemeManager freezes the
	/// application resources as well, so writing to a resource brush throws
	/// InvalidOperationException. UIElement.Opacity is an ordinary dependency
	/// property on a live element and always works.
	///
	/// Semantics: <see cref="Opacity"/> is a ratio against <see cref="DefaultOpacity"/>,
	/// which is the alpha of the old #A0FFFFFF toolbar backdrop (160/255). 1 therefore
	/// reproduces the old look, 0 leaves only the icons.
	/// </summary>
	public static class ToolbarAppearance
	{
		/// <summary>Fully transparent backdrop, only the icons are left.</summary>
		public const double MinimumOpacity = 0.0;

		/// <summary>Completely opaque backdrop.</summary>
		public const double MaximumOpacity = 1.0;

		/// <summary>
		/// Default opacity ratio: 65%.
		///
		/// This used to be 160/255, the alpha of the old hard-coded #A0FFFFFF backdrop,
		/// which is 62.7% and so showed up in the settings box as 62.7 - a number
		/// nobody would ever type on purpose. 65% is the value the design calls for;
		/// writing it literally also keeps the number the slider snaps to and the
		/// number in this constant from drifting apart.
		/// </summary>
		public const double DefaultOpacity = 0.65;

		private static readonly List<WeakReference> Chrome = new List<WeakReference>();

		private static double _opacity = DefaultOpacity;

		/// <summary>
		/// Opacity ratio. 1 is the old look, 0 fades the chrome away completely.
		/// Setting it raises <see cref="OpacityChanged"/>.
		/// </summary>
		public static double Opacity
		{
			get
			{
				return _opacity;
			}
			set
			{
				double clamped = Clamp(value);
				if (Math.Abs(clamped - _opacity) < 0.0001)
				{
					return;
				}
				_opacity = clamped;
				Apply();
				EventHandler handler = OpacityChanged;
				if (handler != null)
				{
					handler(null, EventArgs.Empty);
				}
			}
		}

		/// <summary>Raised after <see cref="Opacity"/> changed, whoever changed it.</summary>
		public static event EventHandler OpacityChanged;

		/// <summary>The same value in percent (0-100), which is what the slider shows.</summary>
		public static double OpacityPercent
		{
			get
			{
				return Math.Round(_opacity * 100.0, 1);
			}
		}

		public static void SetOpacityPercent(double percent)
		{
			ToolbarAppearance.Opacity = percent / 100.0;
		}

		/// <summary>
		/// Registers one piece of chrome (a backdrop Rectangle, a divider, ...) and
		/// immediately gives it the current opacity. Called from the constructors of
		/// the toolbars and the thumbnail panels.
		/// </summary>
		public static void Register(UIElement chrome)
		{
			if (chrome == null)
			{
				return;
			}
			lock (Chrome)
			{
				for (int i = Chrome.Count - 1; i >= 0; i--)
				{
					UIElement existing = Chrome[i].Target as UIElement;
					if (existing == null)
					{
						// the control it belonged to is gone: drop the entry
						Chrome.RemoveAt(i);
					}
					else if (object.ReferenceEquals(existing, chrome))
					{
						return; // already registered
					}
				}
				Chrome.Add(new WeakReference(chrome));
			}
			ApplyTo(chrome);
		}

		/// <summary>Registers several pieces of chrome at once.</summary>
		public static void Register(params UIElement[] chrome)
		{
			if (chrome == null)
			{
				return;
			}
			for (int i = 0; i < chrome.Length; i++)
			{
				ToolbarAppearance.Register(chrome[i]);
			}
		}

		/// <summary>Re-applies the current opacity to every registered piece of chrome.</summary>
		public static void Apply()
		{
			lock (Chrome)
			{
				for (int i = Chrome.Count - 1; i >= 0; i--)
				{
					UIElement element = Chrome[i].Target as UIElement;
					if (element == null)
					{
						Chrome.RemoveAt(i);
						continue;
					}
					ApplyTo(element);
				}
			}
		}

		/// <summary>Forgets every registration.</summary>
		public static void ResetRegistry()
		{
			lock (Chrome)
			{
				Chrome.Clear();
			}
		}

		private static void ApplyTo(UIElement element)
		{
			element.Opacity = _opacity;
		}

		private static double Clamp(double value)
		{
			if (double.IsNaN(value))
			{
				return DefaultOpacity;
			}
			if (value < MinimumOpacity)
			{
				return MinimumOpacity;
			}
			if (value > MaximumOpacity)
			{
				return MaximumOpacity;
			}
			return value;
		}
	}
}
