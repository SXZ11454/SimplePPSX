using System;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Animation;

namespace SimplePPSX
{
	/// <summary>
	/// The rise / sink animation of the toolbars and their tool panels.
	///
	/// Every fold and unfold is pure MOVEMENT. An element that folds starts in its
	/// slot, travels down and leaves through the bottom edge of the window; an element
	/// that unfolds starts below that edge and travels up into its slot. Opacity is
	/// never animated here. A fold that only faded out read as "the toolbar vanished"
	/// instead of "the toolbar went down", which is the whole reason this class
	/// measures the travel distance it does (see <see cref="OffScreenDistanceOf"/>):
	/// once the element is outside the window it does not need to be transparent, it
	/// simply is not drawn.
	///
	/// The rise takes <see cref="EntryDurationSeconds"/> = 0.5 seconds and decelerates
	/// into place (ExponentialEase, ease-out); the sink takes
	/// <see cref="ExitDurationSeconds"/> = 0.3 seconds and accelerates out (CubicEase,
	/// ease-in), because pulling back in feels snappier than leaving. There is no
	/// overshoot and no bounce: an earlier version shot past its slot by 8px and
	/// settled back, which read as a bounce and was not wanted.
	///
	/// Opacity is set, but as a plain value rather than an animation: FloatUp writes
	/// the resting 1.0 before the element can be drawn, which is what undoes the zero
	/// the caller used to hide it and the 30% the idle-dim storyboards leave on the
	/// side toolbars. A running opacity clock is cancelled for the same reason - a
	/// base value cannot be read underneath one.
	///
	/// Both animations are started with BeginAnimation rather than through a
	/// Storyboard, and that is not a style choice. A Storyboard whose target is a
	/// Transform silently does nothing: Storyboard.SetTarget stores the Freezable as
	/// a name reference inside the storyboard, the transform is not registered in the
	/// containing object's namescope, and the lookup fails - with no error. The
	/// opacity part of the same storyboard worked, which is exactly why the toolbars
	/// used to fade in but never move. BeginAnimation targets the object directly and
	/// cannot fail that way.
	///
	/// The thumbnail panels are the one exception: they slide in sideways with a fade
	/// of their own (<see cref="SlideIn"/> / <see cref="SlideOut"/>), because they
	/// stand at the left and right screen edges, where rising from below would read as
	/// a separate box dropping into place rather than as the toolbars opening out.
	/// </summary>
	internal static class ToolbarAnimation
	{
		/// <summary>Length of the rise, in seconds.</summary>
		internal const double EntryDurationSeconds = 0.5;

		/// <summary>Pulling back in is quicker than rising, which feels snappier.</summary>
		private const double ExitDurationSeconds = 0.3;

		/// <summary>
		/// How far past the bottom edge the travel goes, so a pixel of the element
		/// cannot peek over the edge at the start of a rise or the end of a sink.
		/// </summary>
		private const double EdgePad = 2.0;

		/// <summary>Used when the element has no measurable height yet.</summary>
		private const double FallbackDistance = 48.0;

		/// <summary>
		/// Cap on the sideways slide of the thumbnail panels only. The vertical travel
		/// is deliberately not capped: it has to be enough to leave the window.
		/// </summary>
		private const double MaxDistance = 160.0;

		private static readonly Duration EntryDuration = new Duration(TimeSpan.FromSeconds(EntryDurationSeconds));
		private static readonly Duration ExitDuration = new Duration(TimeSpan.FromSeconds(ExitDurationSeconds));

		/// <summary>Returns (and installs if needed) the TranslateTransform that carries the animation.</summary>
		internal static TranslateTransform GetTransform(FrameworkElement element)
		{
			TranslateTransform transform = element.RenderTransform as TranslateTransform;
			if (transform == null)
			{
				transform = new TranslateTransform();
				element.RenderTransform = transform;
			}
			return transform;
		}

		/// <summary>
		/// Brings an element up from below the bottom edge of the window and lets it
		/// come to rest in the slot it already occupies.
		///
		/// It travels the whole of its own height (plus the gap under it), not a
		/// fraction of it, so the movement is the thing that is seen - there is no
		/// fade left to hide a rise that barely moves.
		/// </summary>
		internal static void FloatUp(FrameworkElement element)
		{
			if (element == null)
			{
				return;
			}
			element.Visibility = Visibility.Visible;
			element.IsHitTestVisible = true;

			// Opacity is settled HERE, before the first frame can draw the element.
			// There is no fade to cover for it any more, so the resting value has to be
			// the base value from the moment the element becomes visible - and it has
			// to cancel the clock first, because the idle-dim storyboards leave the
			// side toolbars holding 30% and a base value cannot be read underneath an
			// active animation.
			element.BeginAnimation(UIElement.OpacityProperty, null);
			element.Opacity = 1.0;

			TranslateTransform transform = GetTransform(element);
			Begin(element, transform, EntryDuration, true, null);
		}

		/// <summary>
		/// Slides an element in from the side, instead of up from below.
		///
		/// The thumbnail panels use this rather than <see cref="FloatUp"/>. They are
		/// tall panels standing at the left and right edges of the screen, so rising
		/// reads as a separate box dropping into place; coming in from the screen edge
		/// reads as part of the toolbars opening out. The right-hand panel therefore
		/// travels in the opposite direction to the left-hand one.
		///
		/// Everything else follows FloatUp: same curve, same duration, same treatment
		/// of the resting values. It does keep its fade, and that is deliberate - the
		/// sideways travel is a third of the panel's width rather than a trip off the
		/// screen, so the fade is what carries the last of the move.
		/// <paramref name="fromLeft"/> only chooses the axis and the sign.
		/// </summary>
		internal static void SlideIn(FrameworkElement element, bool fromLeft)
		{
			if (element == null)
			{
				return;
			}

			element.Visibility = Visibility.Visible;
			element.IsHitTestVisible = true;

			TranslateTransform transform = GetTransform(element);
			double distance = HorizontalDistanceOf(element);

			// A panel pinned to the left starts to the left of its slot; one pinned to
			// the right starts to the right of it. Both slide towards the centre.
			double from = fromLeft ? -distance : distance;

			DoubleAnimationUsingKeyFrames slide = new DoubleAnimationUsingKeyFrames
			{
				Duration = EntryDuration
			};
			slide.KeyFrames.Add(new EasingDoubleKeyFrame(from, KeyTime.FromTimeSpan(TimeSpan.Zero)) { EasingFunction = null });
			slide.KeyFrames.Add(new EasingDoubleKeyFrame(0.0, KeyTime.FromTimeSpan(EntryDuration.TimeSpan))
			{
				EasingFunction = new ExponentialEase { Exponent = 5.0, EasingMode = EasingMode.EaseOut }
			});

			DoubleAnimation fade = new DoubleAnimation
			{
				From = 0.0,
				To = 1.0,
				Duration = EntryDuration,
				FillBehavior = FillBehavior.Stop,
				EasingFunction = new CubicEase { EasingMode = EasingMode.EaseOut }
			};

			// The fade is Stop, so the animated value is thrown away the moment it ends
			// and Opacity falls back to whatever base value the caller left - the
			// resting values therefore have to be written here. Begin does the same
			// job for the transform, on the rise's own Completed.
			fade.Completed += delegate
			{
				element.Opacity = 1.0;
				transform.X = 0.0;
				transform.Y = 0.0;
			};

			transform.BeginAnimation(TranslateTransform.XProperty, slide);
			element.BeginAnimation(UIElement.OpacityProperty, fade);
		}

		/// <summary>
		/// Slides an element out towards the side it came from. Collapses it afterwards
		/// through <paramref name="completed"/>, for the same reason FloatDown needs a
		/// caller-supplied collapse: an animation's Completed event is not raised when a
		/// later animation on the same property replaces it.
		/// </summary>
		internal static void SlideOut(FrameworkElement element, bool toLeft, Action completed)
		{
			if (element == null)
			{
				if (completed != null)
				{
					completed();
				}

				return;
			}

			if (element.Visibility == Visibility.Collapsed)
			{
				if (completed != null)
				{
					completed();
				}

				return;
			}

			element.IsHitTestVisible = false;
			TranslateTransform transform = GetTransform(element);
			double distance = HorizontalDistanceOf(element);
			double to = toLeft ? -distance : distance;

			DoubleAnimationUsingKeyFrames slide = new DoubleAnimationUsingKeyFrames { Duration = ExitDuration };
			slide.KeyFrames.Add(new EasingDoubleKeyFrame(0.0, KeyTime.FromTimeSpan(TimeSpan.Zero)) { EasingFunction = null });
			slide.KeyFrames.Add(new EasingDoubleKeyFrame(to, KeyTime.FromTimeSpan(ExitDuration.TimeSpan))
			{
				EasingFunction = new CubicEase { EasingMode = EasingMode.EaseIn }
			});

			DoubleAnimation fade = new DoubleAnimation
			{
				From = 1.0,
				To = 0.0,
				Duration = ExitDuration,
				FillBehavior = FillBehavior.Stop,
				EasingFunction = new CubicEase { EasingMode = EasingMode.EaseOut }
			};

			fade.Completed += delegate
			{
				element.Opacity = 1.0;
				transform.X = 0.0;
				transform.Y = 0.0;

				if (completed != null)
				{
					completed();
				}
			};

			transform.BeginAnimation(TranslateTransform.XProperty, slide);
			element.BeginAnimation(UIElement.OpacityProperty, fade);
		}

		private static double HorizontalDistanceOf(FrameworkElement element)
		{
			// A third of its own width, capped: a full-width slide would take the panel
			// off screen for the whole animation on a wide display.
			double distance = element.ActualWidth / 3.0;
			if (double.IsNaN(distance) || distance <= 0.0)
			{
				return FallbackDistance;
			}

			return distance > MaxDistance ? MaxDistance : distance;
		}

		/// <summary>
		/// Sinks an element down through the bottom edge of the window, then hands
		/// back to <paramref name="completed"/>.
		///
		/// The sink does not touch Opacity, and that is the point: the element stays
		/// exactly as opaque as it was on the way down, so what is seen is the bar
		/// dropping out of sight rather than the bar dimming away. It is hidden by
		/// leaving the window, which does not depend on anything finishing in a
		/// particular order. (Leaving the opacity alone also stops a toolbar the idle
		/// dim had at 30% from flashing bright on its way out.)
		/// </summary>
		internal static void FloatDown(FrameworkElement element, Action completed)
		{
			if (element == null)
			{
				if (completed != null)
				{
					completed();
				}
				return;
			}
			if (element.Visibility == Visibility.Collapsed)
			{
				if (completed != null)
				{
					completed();
				}
				return;
			}

			element.IsHitTestVisible = false;
			TranslateTransform transform = GetTransform(element);
			Begin(element, transform, ExitDuration, false, completed);
		}

		/// <summary>
		/// How far down an element has to travel to be entirely outside the window:
		/// the gap between its bottom edge and the bottom edge of the window, its own
		/// height, whatever it hangs over the top of its own box, and
		/// <see cref="EdgePad"/>.
		///
		/// All of that is measured rather than assumed. The toolbars stand on the
		/// bottom edge, so for them the gap is nothing and the travel is one bar - but
		/// the colour and eraser panels stand ON TOP of the bar, so their gap is the
		/// bar's height and their travel has to include it or they would stop in view.
		/// And the centre toolbar is laid out in a fixed 56px slot while its strip is
		/// as tall as the resolution tier makes it (up to 90), so it hangs over the top
		/// of its host: travel measured to the top of the box would leave that overhang
		/// sitting at the edge for the whole of a rise. Over-travel only lengthens the
		/// move; under-travel leaves a strip of toolbar on screen.
		///
		/// Everything here is device-independent pixels of the same coordinate space
		/// (element measured against the window), so nothing needs converting.
		///
		/// The caller has cleared the element's transform first - Begin does - or this
		/// would be measured from wherever a previous sink left it.
		/// </summary>
		private static double OffScreenDistanceOf(FrameworkElement element)
		{
			double height = element.ActualHeight;
			if (double.IsNaN(height) || height <= 0.0)
			{
				return FallbackDistance;
			}

			double gap = 0.0;
			double overhang = 0.0;
			try
			{
				// How far the element's bottom edge is above the bottom edge of the
				// window it is drawn in.
				FrameworkElement boundary = GetClipBoundary(element);
				if (boundary != null && !double.IsNaN(boundary.ActualHeight) && boundary.ActualHeight > 0.0)
				{
					Point bottom = element.TransformToAncestor(boundary).Transform(new Point(0.0, height));
					gap = boundary.ActualHeight - bottom.Y;

					// Never negative. A boundary shorter than the element's own position
					// means the window has not been laid out at the size it is really
					// going to be yet, and a gap of zero still asks for the full height
					// of travel - which is exactly what an element standing on the edge
					// needs, so the worst case is travelling further, never less.
					if (double.IsNaN(gap) || gap < 0.0)
					{
						gap = 0.0;
					}
				}

				// What the element paints above its own top edge, in its own coordinates.
				Rect descendants = VisualTreeHelper.GetDescendantBounds(element);
				if (!descendants.IsEmpty && descendants.Top < 0.0)
				{
					overhang = -descendants.Top;
				}
			}
			catch (Exception ex)
			{
				// A measurement that throws must not take the fold down with it. Falling
				// back to "stands on the edge" keeps the element out of sight, which is
				// where everything animated here is meant to end up either way.
				gap = 0.0;
				overhang = 0.0;
				Log.WriteDebug("travel measure failed: " + ex.Message);
			}

			double distance = gap + height + overhang + EdgePad;
			if (double.IsNaN(distance) || distance <= 0.0)
			{
				// Already below the edge: a token travel is all that is left to do.
				return FallbackDistance;
			}

			return distance;
		}

		/// <summary>
		/// The element that says where drawing stops: the window itself. Its arranged
		/// box is the rectangle the overlay is clipped to, and measuring to it rather
		/// than to its content also picks up whatever margin the content stands in -
		/// the root grid is inset by a few pixels in one of the adaptive states, and
		/// that inset would otherwise come straight out of the travel and leave a strip
		/// of toolbar at the edge. Returns null when the element is in no window yet.
		/// </summary>
		private static FrameworkElement GetClipBoundary(FrameworkElement element)
		{
			return Window.GetWindow(element);
		}

		private static void Begin(FrameworkElement element, TranslateTransform transform, Duration duration, bool entering, Action completed)
		{
			// Cancel whatever was running and put the element back in its slot: two
			// interrupted animations would otherwise fight over the transform, and the
			// travel distance below has to be measured from the slot rather than from
			// wherever the last one stopped.
			transform.BeginAnimation(TranslateTransform.YProperty, null);
			transform.Y = 0.0;

			double distance = OffScreenDistanceOf(element);
			double from = entering ? distance : 0;
			double to = entering ? 0 : distance;

			// The BASE value, written before the clock starts.
			//
			// On the way in the base is the off-screen position, so a frame drawn
			// before the clock has been sampled puts the element out of sight instead
			// of standing in its slot at full opacity. That frame is not
			// hypothetical: the rise is armed from CompositionTarget.Rendering, and
			// whether the first drawn frame sees the base or the animation is the
			// difference between "comes up from the edge" and "flashes into place".
			//
			// On the way out the base is the resting position, and the sink holds
			// off-screen (HoldEnd, the default) until Completed copies that into the
			// base further down.
			transform.Y = from;

			// One line per travel, because this is the number the whole effect rests on
			// and nobody can see it. travel = box means the element stands on the edge;
			// travel well above it is the gap it has to cross (the panels sit above the
			// bar); travel at the 48px fallback means the measurement failed, which is
			// what "it stopped in view" looks like from the log.
			Log.WriteDebug(string.Format(
				"{0} {1}: travel {2:F1}px (box {3:F1}) over {4:F2}s",
				entering ? "rise" : "sink",
				string.IsNullOrEmpty(element.Name) ? element.GetType().Name : element.Name,
				distance, element.ActualHeight, duration.TimeSpan.TotalSeconds));

			DoubleAnimationUsingKeyFrames rise = new DoubleAnimationUsingKeyFrames();
			rise.Duration = duration;
			rise.KeyFrames.Add(new EasingDoubleKeyFrame(from, KeyTime.FromTimeSpan(TimeSpan.Zero))
			{
				// the value at t=0 only; the easing lives on the second keyframe
				EasingFunction = null
			});

			if (entering)
			{
				// One keyframe pair only: from the off-screen offset up to exactly zero,
				// decelerating. There is deliberately no overshoot - see the note on the
				// class. The value still curves, so it is not a constant-speed slide.
				rise.KeyFrames.Add(new EasingDoubleKeyFrame(to, KeyTime.FromTimeSpan(duration.TimeSpan))
				{
					EasingFunction = new ExponentialEase
					{
						Exponent = 5.0,
						EasingMode = EasingMode.EaseOut
					}
				});
			}
			else
			{
				rise.KeyFrames.Add(new EasingDoubleKeyFrame(to, KeyTime.FromTimeSpan(duration.TimeSpan))
				{
					EasingFunction = new CubicEase
					{
						EasingMode = EasingMode.EaseIn
					}
				});
			}

			// ALWAYS attached, not only when a completion callback was supplied, and
			// attached before the animation starts: WPF does not raise Completed when
			// a later animation replaces this one, so a callback that lives here is a
			// bonus rather than the thing that hides anything.
			//
			// What it does is fold the animated value into the base (a rise would
			// otherwise keep an off-screen base under a finished clock, and a sink a
			// slot) and then hand on to the caller's collapse - the half-bar swap and
			// the panel hide both still want to know when the travel is over.
			//
			// Started on the transform itself, not through a Storyboard: see the note
			// on the class. Opacity is not animated any more, so there is no second
			// clock to keep in step with this one.
			rise.Completed += delegate
			{
				transform.Y = to;

				if (completed != null)
				{
					completed();
				}
			};
			transform.BeginAnimation(TranslateTransform.YProperty, rise);
		}
	}
}
