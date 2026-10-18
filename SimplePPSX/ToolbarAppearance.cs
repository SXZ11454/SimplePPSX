using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows;
using System.Windows.Media;
using System.Windows.Threading;

namespace SimplePPSX
{
	/// <summary>
	/// The opacity control for the presentation toolbars, in two states.
	///
	/// This is the extension point the settings window drives. No toolbar writes a
	/// colour of its own: every piece of chrome (the backdrops, the 1px outlines and
	/// the hairlines between buttons) is a plain Rectangle filled from
	/// ToolbarBackgroundBrush / ToolbarBorderBrush / ToolbarDividerBrush, and the
	/// controls register those Rectangles here. Changing the active opacity then
	/// fades every one of them at once, which means
	///   - the toolbars do not need to know about each other,
	///   - the settings window changes one value and every toolbar follows.
	///
	/// The two states (added together with the second slider in the settings page):
	///   - default: what the backdrop settles on while the pointer is over a
	///     toolbar surface (or a thumbnail panel) - the value the original slider
	///     has always driven, now explicitly "the opacity you get when you are
	///     using the toolbars";
	///   - idle, the "not clicked" state: what the backdrop sinks to the rest of
	///     the time, so a deck is not screened behind its own toolbars. Hovering
	///     any registered surface raises to the default value, leaving glides back
	///     down. The switch itself is a short interpolation, not a jump: a hard
	///     step under the pointer reads as flicker.
	///
	/// Why element Opacity and not brush Opacity: WPF freezes every freezable that
	/// comes out of a resource dictionary, and ModernWpf's ThemeManager freezes the
	/// application resources as well, so writing to a resource brush throws
	/// InvalidOperationException. UIElement.Opacity is an ordinary dependency
	/// property on a live element and always works.
	///
	/// Semantics: <see cref="Opacity"/> and <see cref="IdleOpacity"/> are ratios,
	/// where 1 reproduces the old fully-drawn look and 0 leaves only the icons.
	/// </summary>
	public static class ToolbarAppearance
	{
		/// <summary>Fully transparent backdrop, only the icons are left.</summary>
		public const double MinimumOpacity = 0.0;

		/// <summary>Completely opaque backdrop.</summary>
		public const double MaximumOpacity = 1.0;

		/// <summary>
		/// Default (hover) opacity ratio: 65%.
		///
		/// This used to be 160/255, the alpha of the old hard-coded #A0FFFFFF backdrop,
		/// which is 62.7% and so showed up in the settings box as 62.7 - a number
		/// nobody would ever type on purpose. 65% is the value the design calls for;
		/// writing it literally also keeps the number the slider snaps to and the
		/// number in this constant from drifting apart.
		/// </summary>
		public const double DefaultOpacity = 0.65;

		/// <summary>
		/// Idle ("not clicked") opacity ratio: 35%.
		///
		/// Deliberately below the default so the difference between the two
		/// sliders is visible before the user tunes either of them; the settings
		/// page exposes it, so a taste for more or less presence is one drag away.
		/// </summary>
		public const double DefaultIdleOpacity = 0.35;

		private static readonly List<WeakReference> Chrome = new List<WeakReference>();

		/// <summary>The default-state opacity: what hover settles on.</summary>
		private static double _opacity = DefaultOpacity;

		/// <summary>The idle-state opacity: what everything sinks to when not hovered.</summary>
		private static double _idleOpacity = DefaultIdleOpacity;

		/// <summary>
		/// The opacity the chrome is showing right now, mid-glide included. The
		/// sliders and the hover switch retarget this value; <see cref="Apply"/>
		/// always writes exactly this, so a theme refresh lands on the state the
		/// eye is looking at instead of jumping back to a stale target.
		/// </summary>
		private static double _displayedOpacity = DefaultOpacity;

		/// <summary>
		/// True while the pointer is over a toolbar surface. Starts false: at the
		/// moment a show begins the pointer is over the slide, not the toolbars,
		/// which is precisely the "not clicked" state the idle slider describes.
		/// </summary>
		private static bool _interactive;

		private static DispatcherTimer _glideTimer;

		private static double _glideFrom;

		private static double _glideTo;

		private static readonly Stopwatch _glideClock = new Stopwatch();

		/// <summary>
		/// How long one opacity switch takes. Short enough to keep up with the
		/// pointer, long enough to read as a fade rather than a flicker.
		/// </summary>
		private const double GlideSeconds = 0.18;

		/// <summary>
		/// Opacity ratio of the default (hover) state. 1 is the old look, 0 fades
		/// the chrome away completely. Setting it raises <see cref="OpacityChanged"/>.
		/// </summary>
		public static double Opacity
		{
			get
			{
				return _opacity;
			}
			set
			{
				OnUi(delegate
				{
					SetDefaultExact(value);
				});
			}
		}

		/// <summary>The default-state value in percent (0-100), which is what the slider shows.</summary>
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
		/// Opacity ratio of the idle ("not clicked") state. Interacts with
		/// <see cref="Opacity"/> only through which one is on display.
		/// </summary>
		public static double IdleOpacity
		{
			get
			{
				return _idleOpacity;
			}
			set
			{
				OnUi(delegate
				{
					SetIdleExact(value);
				});
			}
		}

		/// <summary>The idle-state value in percent (0-100), which is what the slider shows.</summary>
		public static double IdleOpacityPercent
		{
			get
			{
				return Math.Round(_idleOpacity * 100.0, 1);
			}
		}

		public static void SetIdleOpacityPercent(double percent)
		{
			ToolbarAppearance.IdleOpacity = percent / 100.0;
		}

		/// <summary>Raised after either state's stored value changed, whoever changed it.</summary>
		public static event EventHandler OpacityChanged;

		/// <summary>
		/// Attaches hover tracking to one toolbar surface. Entering the surface
		/// raises the chrome to <see cref="Opacity"/>, leaving glides it to
		/// <see cref="IdleOpacity"/>. Called from the constructors of the toolbars
		/// and the thumbnail panels, next to their <see cref="Register"/> calls.
		/// </summary>
		public static void HookHover(UIElement surface)
		{
			if (surface == null)
			{
				return;
			}
			surface.MouseEnter += delegate
			{
				ToolbarAppearance.SetInteractive(true);
			};
			surface.MouseLeave += delegate
			{
				ToolbarAppearance.SetInteractive(false);
			};
		}

		/// <summary>
		/// Switches between the two opacity states. A no-op when the state did not
		/// change, so overlapping surfaces (thumbnails sliding over a toolbar)
		/// renegotiating hover never restart a glide that is already aimed right.
		/// </summary>
		public static void SetInteractive(bool interactive)
		{
			OnUi(delegate
			{
				if (_interactive == interactive)
				{
					return;
				}
				_interactive = interactive;
				Retarget();
			});
		}

		/// <summary>
		/// Registers one piece of chrome (a backdrop Rectangle, a divider, ...) and
		/// immediately gives it the currently displayed opacity. Called from the
		/// constructors of the toolbars and the thumbnail panels.
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

		/// <summary>Re-applies the displayed opacity to every registered piece of chrome.</summary>
		public static void Apply()
		{
			OnUi(delegate
			{
				ApplyCore();
			});
		}

		/// <summary>Forgets every registration.</summary>
		public static void ResetRegistry()
		{
			lock (Chrome)
			{
				Chrome.Clear();
			}
		}

		private static void ApplyCore()
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

		private static void ApplyTo(UIElement element)
		{
			element.Opacity = _displayedOpacity;
		}

		private static void SetDefaultExact(double value)
		{
			double clamped = Clamp(value);
			if (Math.Abs(clamped - _opacity) < 0.0001)
			{
				return;
			}
			_opacity = clamped;
			if (_interactive)
			{
				Retarget();
			}
			EventHandler handler = OpacityChanged;
			if (handler != null)
			{
				handler(null, EventArgs.Empty);
			}
		}

		private static void SetIdleExact(double value)
		{
			double clamped = Clamp(value);
			if (Math.Abs(clamped - _idleOpacity) < 0.0001)
			{
				return;
			}
			_idleOpacity = clamped;
			if (!_interactive)
			{
				Retarget();
			}
			EventHandler handler = OpacityChanged;
			if (handler != null)
			{
				handler(null, EventArgs.Empty);
			}
		}

		/// <summary>Aims the displayed opacity at whatever the current state values.</summary>
		private static void Retarget()
		{
			StartGlide(_interactive ? _opacity : _idleOpacity);
		}

		/// <summary>
		/// The one switch between two opacities, interpolated over
		/// <see cref="GlideSeconds"/> with a cubic ease-out - the same shape as the
		/// toolbar plate and ring movements, so chrome and motion read as one
		/// design. Equal targets skip the timer entirely.
		/// </summary>
		private static void StartGlide(double target)
		{
			if (double.IsNaN(_displayedOpacity))
			{
				_displayedOpacity = target;
				ApplyCore();
				return;
			}
			if (Math.Abs(target - _displayedOpacity) < 0.0001)
			{
				StopGlide();
				_displayedOpacity = target;
				ApplyCore();
				return;
			}
			_glideFrom = _displayedOpacity;
			_glideTo = target;
			_glideClock.Reset();
			_glideClock.Start();
			if (_glideTimer == null)
			{
				_glideTimer = new DispatcherTimer
				{
					Interval = TimeSpan.FromMilliseconds(16.0)
				};
				_glideTimer.Tick += ToolbarAppearance.GlideTick;
			}
			_glideTimer.Start();
		}

		private static void GlideTick(object sender, EventArgs e)
		{
			double t = _glideClock.Elapsed.TotalSeconds / GlideSeconds;
			if (t >= 1.0)
			{
				StopGlide();
				_displayedOpacity = _glideTo;
				ApplyCore();
				return;
			}
			double eased = 1.0 - Math.Pow(1.0 - t, 3.0);
			_displayedOpacity = _glideFrom + (_glideTo - _glideFrom) * eased;
			ApplyCore();
		}

		private static void StopGlide()
		{
			if (_glideTimer != null)
			{
				_glideTimer.Stop();
			}
			_glideClock.Reset();
		}

		/// <summary>
		/// Runs an action on the UI dispatcher: inline when the caller already is
		/// there (which is every hover event and settings write in this app), posted
		/// otherwise. Without this a stray non-UI caller would touch WPF state.
		/// </summary>
		private static void OnUi(Action action)
		{
			Dispatcher dispatcher = (Application.Current != null) ? Application.Current.Dispatcher : null;
			if (dispatcher == null || dispatcher.CheckAccess())
			{
				action();
				return;
			}
			dispatcher.BeginInvoke(action);
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
