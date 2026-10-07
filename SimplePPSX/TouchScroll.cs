using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;

namespace SimplePPSX
{
	/// <summary>
	/// Finger scrolling for every scroll container.
	///
	/// Copied from SimpleRollCall's TouchScroll, with the same behaviour and the same
	/// deliberate choices: dragging over the content scrolls it, with a short glide
	/// once the finger lifts. It is copied rather than referenced because the two
	/// applications share no code and the alternative - a mouse wheel and a thin
	/// scrollbar - is unusable on the touch panels this is built for.
	///
	/// Attach <c>spx:TouchScroll.Enabled="True"</c>.
	///
	/// The deliberate details, kept identical to the original:
	///   - the gesture only starts once the finger has moved away from its start
	///     point, so taps on buttons, rows and list items keep working. That matters
	///     most here, where a tap selects a slide and a drag must scroll past it;
	///   - the scroll view inside a text box keeps out of it: dragging there has to
	///     keep selecting text;
	///   - a second finger is ignored - no pinch, one finger scrolls;
	///   - an item-scrolling view counts items rather than pixels, so the finger
	///     distance is converted with the size of one visible item.
	/// </summary>
	internal static class TouchScroll
	{
		/// <summary>Attached property that turns finger scrolling on for a scroll view.</summary>
		public static readonly DependencyProperty EnabledProperty = DependencyProperty.RegisterAttached(
			"Enabled",
			typeof(bool),
			typeof(TouchScroll),
			new PropertyMetadata(false, OnEnabledChanged));

		public static bool GetEnabled(DependencyObject obj)
		{
			return (bool)obj.GetValue(EnabledProperty);
		}

		public static void SetEnabled(DependencyObject obj, bool value)
		{
			obj.SetValue(EnabledProperty, value);
		}

		/// <summary>Finger travel, in device independent pixels, before a touch becomes a scroll.</summary>
		private const double DragThreshold = 8.0;

		/// <summary>Slowest fling (pixels per millisecond) still worth gliding on.</summary>
		private const double MinFlingSpeed = 0.15;

		/// <summary>Share of the fling speed left after one second; 0.05 stops it quickly.</summary>
		private const double FlingDecayPerSecond = 0.05;

		/// <summary>One glide tick per display frame.</summary>
		private static readonly TimeSpan TickInterval = TimeSpan.FromMilliseconds(16);

		private static readonly DependencyProperty StateProperty = DependencyProperty.RegisterAttached(
			"State",
			typeof(ScrollState),
			typeof(TouchScroll),
			new PropertyMetadata(null));

		/// <summary>One gesture at a time per scroll viewer. UI thread only.</summary>
		private sealed class ScrollState
		{
			public TouchDevice Touch;
			public Point Start;
			public Point Last;
			public DateTime LastTime;
			public double VelocityX;
			public double VelocityY;
			public bool Dragging;
			public bool ScrollAllowed = true;
			public DispatcherTimer Glide;
		}

		private static void OnEnabledChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
		{
			ScrollViewer viewer = d as ScrollViewer;
			if (viewer == null)
			{
				return;
			}

			if ((bool)e.NewValue)
			{
				viewer.PreviewTouchDown += OnTouchDown;
				viewer.PreviewTouchMove += OnTouchMove;
				viewer.PreviewTouchUp += OnTouchUp;
				viewer.LostTouchCapture += OnLostTouchCapture;
				viewer.Unloaded += OnUnloaded;
				return;
			}

			viewer.PreviewTouchDown -= OnTouchDown;
			viewer.PreviewTouchMove -= OnTouchMove;
			viewer.PreviewTouchUp -= OnTouchUp;
			viewer.LostTouchCapture -= OnLostTouchCapture;
			viewer.Unloaded -= OnUnloaded;

			ScrollState state = GetState(viewer);
			StopGlide(state);
			if (state.Touch != null)
			{
				state.Touch.Capture(null);
				state.Touch = null;
			}
		}

		/// <summary>The finger position in the scroll view's own coordinates.</summary>
		private static Point PositionOf(TouchEventArgs e, ScrollViewer viewer)
		{
			return e.GetTouchPoint(viewer).Position;
		}

		private static ScrollState GetState(ScrollViewer viewer)
		{
			ScrollState state = viewer.GetValue(StateProperty) as ScrollState;
			if (state == null)
			{
				state = new ScrollState();
				viewer.SetValue(StateProperty, state);
			}

			return state;
		}

		// ================= gesture =================

		private static void OnTouchDown(object sender, TouchEventArgs e)
		{
			ScrollViewer viewer = (ScrollViewer)sender;
			ScrollState state = GetState(viewer);

			// A second finger is not part of the gesture and never stops the first.
			if (state.Touch != null)
			{
				return;
			}

			// Touching the content also cancels a running fling.
			StopGlide(state);
			state.Touch = e.TouchDevice;
			state.Start = PositionOf(e, viewer);
			state.Last = state.Start;
			state.LastTime = DateTime.UtcNow;
			state.VelocityX = 0.0;
			state.VelocityY = 0.0;
			state.Dragging = false;

			// The scroll view inside a text box carries the caret and the selection:
			// scrolling there would steal both.
			state.ScrollAllowed = FindAncestor<TextBoxBase>(viewer) == null;
		}

		private static void OnTouchMove(object sender, TouchEventArgs e)
		{
			ScrollViewer viewer = (ScrollViewer)sender;
			ScrollState state = GetState(viewer);

			if (!state.ScrollAllowed || !object.ReferenceEquals(state.Touch, e.TouchDevice))
			{
				return;
			}

			Point point = PositionOf(e, viewer);
			DateTime now = DateTime.UtcNow;

			if (!state.Dragging)
			{
				if (Math.Abs(point.X - state.Start.X) < DragThreshold &&
					Math.Abs(point.Y - state.Start.Y) < DragThreshold)
				{
					// Still a plain tap.
					return;
				}

				// Past the threshold the gesture becomes a scroll. The finger is claimed
				// so it keeps arriving here even after the content has moved out from
				// under it, and any mouse capture a pressed button below the finger took
				// is dropped - otherwise lifting the finger would still click it.
				state.Dragging = true;
				Mouse.Captured?.ReleaseMouseCapture();
				state.Touch.Capture(viewer);
				e.Handled = true;
			}

			double milliseconds = Math.Max((now - state.LastTime).TotalMilliseconds, 1.0);
			double dx = point.X - state.Last.X;
			double dy = point.Y - state.Last.Y;

			ScrollBy(viewer, dx, dy);

			// Smoothed, so a shaky finger does not shoot the list off.
			state.VelocityX = (0.65 * (dx / milliseconds)) + (0.35 * state.VelocityX);
			state.VelocityY = (0.65 * (dy / milliseconds)) + (0.35 * state.VelocityY);
			state.Last = point;
			state.LastTime = now;
			e.Handled = true;
		}

		private static void OnTouchUp(object sender, TouchEventArgs e)
		{
			ScrollViewer viewer = (ScrollViewer)sender;
			ScrollState state = GetState(viewer);

			if (!object.ReferenceEquals(state.Touch, e.TouchDevice))
			{
				return;
			}

			bool wasDragging = state.Dragging;
			state.Touch = null;
			state.Dragging = false;

			// Raises LostTouchCapture, which now sees no gesture in progress.
			e.TouchDevice.Capture(null);

			if (wasDragging)
			{
				// A scroll never ends in a click.
				e.Handled = true;
				StartGlide(viewer, state);
			}
		}

		/// <summary>
		/// The finger went away without an up event - capture stolen, element replaced.
		/// End the gesture but let the fling carry on, as a normal release would.
		/// </summary>
		private static void OnLostTouchCapture(object sender, TouchEventArgs e)
		{
			ScrollViewer viewer = (ScrollViewer)sender;
			ScrollState state = GetState(viewer);

			if (!object.ReferenceEquals(state.Touch, e.TouchDevice))
			{
				return;
			}

			bool wasDragging = state.Dragging;
			state.Touch = null;
			state.Dragging = false;

			if (wasDragging)
			{
				StartGlide(viewer, state);
			}
		}

		/// <summary>A scroll view leaving the tree stops everything.</summary>
		private static void OnUnloaded(object sender, RoutedEventArgs e)
		{
			ScrollState state = GetState((ScrollViewer)sender);
			StopGlide(state);
			if (state.Touch != null)
			{
				state.Touch.Capture(null);
				state.Touch = null;
			}

			state.Dragging = false;
		}

		// ================= fling =================

		private static void StartGlide(ScrollViewer viewer, ScrollState state)
		{
			if (Math.Abs(state.VelocityX) < MinFlingSpeed && Math.Abs(state.VelocityY) < MinFlingSpeed)
			{
				return;
			}

			DispatcherTimer timer = new DispatcherTimer(DispatcherPriority.Background) { Interval = TickInterval };
			DateTime previous = DateTime.UtcNow;
			state.Glide = timer;

			timer.Tick += delegate
			{
				DateTime now = DateTime.UtcNow;

				// A stalled frame must not teleport the list.
				double milliseconds = Math.Min((now - previous).TotalMilliseconds, 100.0);
				previous = now;

				double decay = Math.Pow(FlingDecayPerSecond, milliseconds / 1000.0);
				state.VelocityX *= decay;
				state.VelocityY *= decay;

				double beforeX = viewer.HorizontalOffset;
				double beforeY = viewer.VerticalOffset;
				ScrollBy(viewer, state.VelocityX * milliseconds, state.VelocityY * milliseconds);

				bool stuck = (Math.Abs(viewer.HorizontalOffset - beforeX) < 0.01)
					&& (Math.Abs(viewer.VerticalOffset - beforeY) < 0.01);
				bool slow = Math.Abs(state.VelocityX) < MinFlingSpeed
					&& Math.Abs(state.VelocityY) < MinFlingSpeed;

				if (stuck || slow || !viewer.IsLoaded)
				{
					StopGlide(state);
				}
			};

			timer.Start();
		}

		private static void StopGlide(ScrollState state)
		{
			DispatcherTimer timer = state.Glide;
			if (timer != null)
			{
				state.Glide = null;
				timer.Stop();
			}

			state.VelocityX = 0.0;
			state.VelocityY = 0.0;
		}

		// ================= scrolling =================

		/// <summary>
		/// Moves the content with the finger: dragging down brings the part above into
		/// view, so the offsets change against the finger distance.
		/// </summary>
		private static void ScrollBy(ScrollViewer viewer, double dx, double dy)
		{
			if (dx != 0.0 && viewer.ScrollableWidth > 0.0)
			{
				viewer.ScrollToHorizontalOffset(viewer.HorizontalOffset - ToUnits(viewer, dx, true));
			}

			if (dy != 0.0 && viewer.ScrollableHeight > 0.0)
			{
				viewer.ScrollToVerticalOffset(viewer.VerticalOffset - ToUnits(viewer, dy, false));
			}
		}

		/// <summary>
		/// A pixel scroll view counts pixels and an item scroll view counts items, so a
		/// finger distance is turned into the view's own units: the same share of the
		/// viewport, expressed as pixels * visible items / viewport pixels.
		/// </summary>
		private static double ToUnits(ScrollViewer viewer, double pixels, bool horizontal)
		{
			if (!viewer.CanContentScroll)
			{
				return pixels;
			}

			ScrollContentPresenter presenter = viewer.Template?.FindName("PART_ScrollContentPresenter", viewer)
				as ScrollContentPresenter;

			double screen;
			if (presenter == null)
			{
				screen = horizontal ? viewer.ActualWidth : viewer.ActualHeight;
			}
			else
			{
				screen = horizontal ? presenter.ActualWidth : presenter.ActualHeight;
			}

			double items = horizontal ? viewer.ViewportWidth : viewer.ViewportHeight;
			if (screen <= 0.0 || items <= 0.0)
			{
				return pixels;
			}

			return pixels * items / screen;
		}

		// ================= tree helpers =================

		private static T FindAncestor<T>(DependencyObject element) where T : DependencyObject
		{
			DependencyObject current = element;
			while (current != null)
			{
				T match = current as T;
				if (match != null)
				{
					return match;
				}

				current = GetParent(current);
			}

			return null;
		}

		private static DependencyObject GetParent(DependencyObject obj)
		{
			if (obj is Visual || obj is System.Windows.Media.Media3D.Visual3D)
			{
				return VisualTreeHelper.GetParent(obj);
			}

			FrameworkContentElement content = obj as FrameworkContentElement;
			return content != null ? content.Parent : LogicalTreeHelper.GetParent(obj);
		}
	}
}