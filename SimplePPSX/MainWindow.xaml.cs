using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using Microsoft.Office.Interop.PowerPoint;

namespace SimplePPSX
{
	/// <summary>
	/// The transparent full-screen overlay that floats above PowerPoint while a
	/// presentation is running: three toolbars plus two thumbnail panels.
	///
	/// Division of labour:
	///   - MainWindow.xaml only declares the z-order and the empty hosting Grids,
	///   - the three toolbars are pinned to the bottom edge, left / center / right
	///     (ApplyToolbarPlacements),
	///   - the rise / sink comes from <see cref="ToolbarAnimation"/> (0.5s,
	///     non-linear),
	///   - the backdrop colours come from ToolbarBackgroundBrush, which is the very
	///     brush the opacity setting drives.
	/// </summary>
	public partial class MainWindow : Window
	{
		private bool IsTransparent
		{
			get
			{
				return base.Opacity < 0.1;
			}
		}

		/// <summary>Set while exiting: the timer stops touching COM.</summary>
		public bool IsExiting { get; set; }

		/// <summary>
		/// Never let this window actually close.
		///
		/// MainWindow is built ONCE, in ViewHelper.BuildViews, and reused for every
		/// presentation. A Window that has been closed cannot be shown again, so any
		/// close - by Alt+F4, by a stray Close(), by a style that sets CanClose - ends
		/// the toolbar overlay for the rest of the session: the next presentation builds
		/// no views (BuildViews returns early because MainWindow is not null) and no
		/// toolbar ever appears again, with no way to bring it back short of restarting
		/// the app. That is exactly "Alt+F4 closes it and it never comes back".
		///
		/// SetAltF4Enabled(false) was supposed to prevent this by swallowing the key,
		/// but it only does so while this window has keyboard focus - and the overlay is
		/// a click-through window that normally does NOT have focus, since PowerPoint's
		/// slideshow holds it. So Alt+F4 reached Windows, which closed the window.
		///
		/// Cancelling the close is the reliable place to stop it, because it does not
		/// depend on focus: the window is only ever hidden here, never destroyed, and
		/// ShowBegin can always bring it back. The one legitimate close is the tray's
		/// Exit, which sets IsExiting first.
		/// </summary>
		protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
		{
			if (!this.IsExiting)
			{
				e.Cancel = true;
				this.HiddenWindow();
				return;
			}

			base.OnClosing(e);
		}

		public DispatcherTimer UpdateTimer { get; set; }

		public bool IsNeedUpdateTimer { get; set; } = true;

		/// <summary>Is the left thumbnail panel open?</summary>
		public bool IsLeftSlideManagerViewShowed { get; set; }

		/// <summary>Is the right thumbnail panel open?</summary>
		public bool IsRightSlideManagerViewShowed { get; set; }

		/// <summary>Left thumbnail panel, created in MainWindow_Loaded.</summary>
		public SlideThumbnailPanel LeftThumbnailPanel { get; private set; }

		/// <summary>Right thumbnail panel, created in MainWindow_Loaded.</summary>
		public SlideThumbnailPanel RightThumbnailPanel { get; private set; }

		/// <summary>
		/// True once MainWindow_Loaded has run - that is, once the window itself is
		/// built. Only used for diagnostics; what decides whether show-begin reaches
		/// this window is IsShowBeginSubscribed.
		/// </summary>
		internal bool IsWindowLoaded
		{
			get
			{
				return this._isWindowLoaded;
			}
		}

		/// <summary>
		/// True once MainWindow_Loaded has run. Loaded fires again after Hide() - the
		/// Alt+F4 path hides the window rather than closing it - and everything in
		/// there is one-time: running it twice would add the toolbars twice and
		/// subscribe ShowBegin twice, which is how show-begin ends up executing twice.
		/// </summary>
		private bool _isWindowLoaded;

		/// <summary>Have the five toolbars been put into their hosts yet?</summary>
		private bool _toolbarsAttached;

		/// <summary>
		/// Are the four side toolbars folded away with the centre bar? A fold never
		/// survives a presentation; see _pptOperatorShowBegin.
		/// </summary>
		private bool _isSideToolbarsFolded;

		/// <summary>
		/// Safety net for the entry animation. See <see cref="PlayEntryAnimation"/>.
		/// </summary>
		private DispatcherTimer _entryFrameTimer;

		/// <summary>The frame handler the entry animation is waiting on, if it has one.</summary>
		private EventHandler _entryRenderingHandler;

		/// <summary>The one-shot hold before the analysis runs. See StartAnalysisAfterEntry.</summary>
		private DispatcherTimer _analysisDelayTimer;

		public MainWindow()
		{
			this.InitializeComponent();
			base.DataContext = this;
			// Keep watching the theme for as long as this window lives, so a light/dark
			// switch re-squares the corner radii and re-registers the toolbar brushes.
			ThemeRadiusFix.Watch(this);
			this.UpdateTimer = new DispatcherTimer
			{
				Interval = TimeSpan.FromMilliseconds(MainWindow.TimerSpan)
			};
			this.UpdateTimer.Tick += this.UpdateTimerTick;
			this._releaseTimer = new DispatcherTimer
			{
				Interval = TimeSpan.FromMilliseconds(MainWindow.ReleaseTimerSpan)
			};
			this._releaseTimer.Tick += this.ReleaseTimerTick;
			base.Loaded += this.MainWindow_Loaded;
			base.PreviewKeyDown += this.MainWindow_KeyDown;
			base.SizeChanged += this.MainWindow_SizeChanged;
			base.LocationChanged += this.MainWindow_LocationChanged;
			base.ManipulationBoundaryFeedback += delegate(object s, ManipulationBoundaryFeedbackEventArgs args)
			{
				args.Handled = true;
			};

			// A thumbnail panel is as wide as the side toolbar it stands on, and that
			// width is measured rather than written down - so a toolbar that changes
			// size has to re-run the placement or the panels keep the width they
			// happened to have.
			//
			// This cannot be left to MainWindow_SizeChanged. At the first layout the
			// toolbars are still Collapsed - the entry animation is what reveals them -
			// so all four hosts measure 0 and both panel widths come out as NaN. The
			// toolbars' own SizeChanged is the only event that fires once they have
			// actually been measured.
			FrameworkElement[] toolbarHosts =
			{
				this.LeftToolbarHost,
				this.LeftBottomToolbarHost,
				this.RightToolbarHost,
				this.RightBottomToolbarHost
			};
			foreach (FrameworkElement host in toolbarHosts)
			{
				host.SizeChanged += delegate
				{
					this.ApplyToolbarPlacements();
				};
			}

			// The COM subscriptions live here rather than in MainWindow_Loaded. What
			// decides whether show-begin reaches this window by itself is this - not
			// Loaded, which is a side effect of Show() and can fail or be delayed by
			// the preload switch. ViewHelper.Instance_ShowBegin reads
			// IsShowBeginSubscribed to know whether to call in directly, and getting
			// that wrong either way costs the whole show: two runs of show-begin
			// (two entry animations armed, two analysis passes queued) or none.
			PptOperator.Instance.ShowBegin += this._pptOperatorShowBegin;
			PptOperator.Instance.ShowEnd += this._pptOperatorShowEnd;
			PptOperator.Instance.NextSlide += this.PptOperatorNextSlide;
			this.IsShowBeginSubscribed = true;
		}

		/// <summary>
		/// True once the COM events above are wired: from then on show-begin reaches
		/// this window through the event itself. Read by ViewHelper.Instance_ShowBegin.
		/// </summary>
		internal bool IsShowBeginSubscribed { get; private set; }

		private void MainWindow_Loaded(object sender, RoutedEventArgs e)
		{
			if (this._isWindowLoaded)
			{
				return;
			}
			this._isWindowLoaded = true;

			// Both thumbnail panels, both bound to the same SlideItems
			this.LeftThumbnailPanel = new SlideThumbnailPanel();
			this.RightThumbnailPanel = new SlideThumbnailPanel();
			this.LeftThumbnailHost.Children.Add(this.LeftThumbnailPanel);
			this.RightThumbnailHost.Children.Add(this.RightThumbnailPanel);
			this.LeftThumbnailPanel.SlideListBox.ManipulationBoundaryFeedback += delegate(object s, ManipulationBoundaryFeedbackEventArgs args)
			{
				args.Handled = true;
			};
			this.RightThumbnailPanel.SlideListBox.ManipulationBoundaryFeedback += delegate(object s, ManipulationBoundaryFeedbackEventArgs args)
			{
				args.Handled = true;
			};

			// A click on a thumbnail, or on the number above it, moves the show. The
			// panels do the hit testing; MainWindow does the moving, because it is the
			// only place that knows how a slide is addressed in the running show (a
			// custom show, a slide range, a starting offset).
			this.LeftThumbnailPanel.SlideRequested += this.ThumbnailPanel_SlideRequested;
			this.RightThumbnailPanel.SlideRequested += this.ThumbnailPanel_SlideRequested;

			// The entries are sized from what the panel's ScrollViewer leaves over, and
			// the panel works that out for itself - it is the only thing that can see
			// its own scrollbar. This is the only way the answer reaches us: the panel
			// is as wide as the toolbar beside it, and neither of those changes width
			// when the scrollbar takes its share, so nothing else would fire.
			this.LeftThumbnailPanel.ContentWidthChanged += this.ThumbnailPanel_ContentWidthChanged;
			this.RightThumbnailPanel.ContentWidthChanged += this.ThumbnailPanel_ContentWidthChanged;

			// The panels sit above the click-through mask now, so they have to close
			// themselves when the press lands on their backing. See DismissRequested.
			SlideThumbnailPanel leftPanel = this.LeftThumbnailPanel;
			this.LeftThumbnailPanel.DismissRequested += delegate
			{
				this.HideThumbnailPanel(leftPanel);
			};
			SlideThumbnailPanel rightPanel = this.RightThumbnailPanel;
			this.RightThumbnailPanel.DismissRequested += delegate
			{
				this.HideThumbnailPanel(rightPanel);
			};

			// Office 2016 and newer get the horizontal toolbars, older versions the
			// vertical ones. Either way the row sits along the bottom edge
			// (ApplyToolbarPlacements lays out left / center / right).
			//
			// The attachment itself lives in EnsureToolbarsAttached, not here: with the
			// preload switch on, this window loads at start-up, when the presentation
			// app is not attached yet and IsOffice2016OrLater is therefore still false.
			// Building the wrong variant here and never rebuilding it is exactly what
			// that would cost, so the toolbars are only attached once the version is
			// known (or at show-begin, which is where this always used to run).
			if (PptVersionManager.SoftWareType != null)
			{
				this.EnsureToolbarsAttached();
			}

			// Both tool panels belong to BottomArea (it owns their state and their slider)
			// but are hosted above the toolbar row, so showing one does not push the
			// toolbars around. Only one is ever up, so the stack is one tall at most.
			if (ViewHelper.BottomArea.EraserPanel != null)
			{
				this.EraserPanelHost.Children.Add(ViewHelper.BottomArea.EraserPanel);
			}
			if (ViewHelper.BottomArea.ColorPanel != null)
			{
				this.ColorPanelHost.Children.Add(ViewHelper.BottomArea.ColorPanel);
			}

			base.Width = SystemParameters.PrimaryScreenWidth;
			base.Height = SystemParameters.PrimaryScreenHeight;

			this.ApplyToolbarPlacements();
			this.AdjustFontSizeAndStyle();
			this.HookToolMenuDismissal();
			this.SubscribeSelectionChangedEvent();
			this._gesture = new Gesture();
			this._gesture.ChangeGestureScope();
			WindowsApi.SetPopupParameter();

			// No entry animation here any more. Loaded runs at start-up when the window
			// is preloaded, and a rise played then would be spent on a window nobody can
			// see - the toolbars would simply be there, at rest, by the time the show
			// starts. _pptOperatorShowBegin is what plays it now, once per presentation,
			// so the animation is always spent on the moment it is supposed to be seen.
		}

		/// <summary>
		/// Puts the five toolbar hosts' children in place. Idempotent, because the
		/// window can load before the presentation app is attached (see
		/// MainWindow_Loaded) and show-begin must not add a second copy.
		/// </summary>
		private void EnsureToolbarsAttached()
		{
			if (this._toolbarsAttached)
			{
				return;
			}
			this._toolbarsAttached = true;

			if (PptVersionManager.IsOffice2016OrLater)
			{
				this.LeftBottomToolbarHost.Children.Add(ViewHelper.LeftBottomArea);
				this.RightBottomToolbarHost.Children.Add(ViewHelper.RightBottomArea);
			}
			else
			{
				this.LeftToolbarHost.Children.Add(ViewHelper.LeftArea);
				this.RightToolbarHost.Children.Add(ViewHelper.RightArea);
			}
			this.CenterToolbarHost.Children.Add(ViewHelper.BottomArea);
		}

		/// <summary>
		/// The entry animation: left, center and right toolbar rise together
		/// over 0.5 seconds, with the thumbnail panels one layer below them (they play
		/// the same animation when the user opens them).
		///
		/// Played from _pptOperatorShowBegin on every presentation, not from Loaded -
		/// a preloaded window loads once at start-up, long before there is anything to
		/// animate onto.
		///
		/// The rise does not start here. It starts on the first frame that actually
		/// draws the toolbars - see <see cref="StartEntryAnimationOnFirstFrame"/>.
		/// </summary>
		private void PlayEntryAnimation()
		{
			// Put every host on screen now, but at zero opacity, so it is laid out and
			// painted (an invisible element still takes part in layout) while the rise
			// waits. Nothing is showing for this one frame: that is the point.
			this.SetEntryHostsPrepared(true);

			this.StartEntryAnimationOnFirstFrame();

			// The thumbnail panels start closed, so they are not part of the entry
			// animation; see ToggleThumbnailPanel for when they play it.
		}

		/// <summary>Visibility and zero opacity for the five toolbar hosts.</summary>
		private void SetEntryHostsPrepared(bool prepared)
		{
			// The opacity clock goes first, because these are base values and an active
			// clock wins over them: the idle-dim storyboards leave the side toolbars
			// holding 30%, and prepared has to make them invisible. Anything left over
			// from the last presentation - a fold's sink, a rise still in flight - is a
			// TRANSFORM, and it cannot show anything while the opacity reads zero;
			// FloatUp cancels it when the rise actually starts.
			this.LeftToolbarHost.BeginAnimation(UIElement.OpacityProperty, null);
			this.LeftBottomToolbarHost.BeginAnimation(UIElement.OpacityProperty, null);
			this.CenterToolbarHost.BeginAnimation(UIElement.OpacityProperty, null);
			this.RightToolbarHost.BeginAnimation(UIElement.OpacityProperty, null);
			this.RightBottomToolbarHost.BeginAnimation(UIElement.OpacityProperty, null);

			this.SetToolbarsOpacity(prepared ? 0.0 : 1.0);
			this.LeftToolbarHost.Visibility = Visibility.Visible;
			this.LeftBottomToolbarHost.Visibility = Visibility.Visible;
			this.CenterToolbarHost.Visibility = Visibility.Visible;
			this.RightToolbarHost.Visibility = Visibility.Visible;
			this.RightBottomToolbarHost.Visibility = Visibility.Visible;
		}

		/// <summary>
		/// Starts the rise on the first rendered frame that contains the toolbars.
		///
		/// Why not straight away: a WPF clock counts in wall-clock time, not in frames.
		/// The show-begin path is busy - the toolbars are built, the slide is exported,
		/// the window makes its first present - and any of that can sit on the UI thread
		/// for longer than the whole 0.5 second rise. The clock then runs out while not
		/// one frame of it has been drawn, and all the user ever sees is the toolbar
		/// standing still, which is exactly the "no rise animation" report.
		///
		/// Waiting for a frame that already contains the toolbars means the first thing
		/// the animation does is what gets drawn. CompositionTarget.Rendering fires once
		/// per presented frame, so the handler is removed on the first call; the timer is
		/// only there in case no frame comes (the window is covered, or the show ended
		/// first), so the toolbar cannot be left invisible at zero opacity.
		/// </summary>
		private void StartEntryAnimationOnFirstFrame()
		{
			// Two shows within one frame interval would otherwise arm this twice, and
			// both handlers would rise the toolbars - the second one restarting the
			// animation the first had just begun.
			this.StopEntryRendering();
			this.StopEntryFrameTimer();

			bool started = false;
			EventHandler onRendering = null;

			onRendering = delegate (object sender, EventArgs e)
			{
				if (started)
				{
					return;
				}
				started = true;
				this.StopEntryRendering();
				this.StopEntryFrameTimer();
				Log.WriteDebug("entry rise: first rendered frame");
				this.RaiseToolbars();
			};
			this._entryRenderingHandler = onRendering;
			CompositionTarget.Rendering += onRendering;

			DispatcherTimer timer = new DispatcherTimer
			{
				Interval = TimeSpan.FromMilliseconds(MainWindow.EntryAnimationFallbackMilliseconds)
			};
			timer.Tick += delegate (object sender, EventArgs e)
			{
				if (started)
				{
					return;
				}
				started = true;
				this.StopEntryRendering();
				this.StopEntryFrameTimer();
				Log.WriteDebug(string.Format(
					"entry rise: NO frame within {0}ms, forcing the rise",
					MainWindow.EntryAnimationFallbackMilliseconds));
				this.RaiseToolbars();
			};
			this._entryFrameTimer = timer;
			timer.Start();
		}

		private void StopEntryRendering()
		{
			if (this._entryRenderingHandler != null)
			{
				CompositionTarget.Rendering -= this._entryRenderingHandler;
				this._entryRenderingHandler = null;
			}
		}

		private void StopEntryFrameTimer()
		{
			if (this._entryFrameTimer != null)
			{
				this._entryFrameTimer.Stop();
				this._entryFrameTimer = null;
			}
		}

		/// <summary>The five hosts rise together. The thumbnail panels play their own.</summary>
		private void RaiseToolbars()
		{
			this.SetEntryHostsPrepared(false);
			ToolbarAnimation.FloatUp(this.LeftToolbarHost);
			ToolbarAnimation.FloatUp(this.LeftBottomToolbarHost);
			ToolbarAnimation.FloatUp(this.CenterToolbarHost);
			ToolbarAnimation.FloatUp(this.RightToolbarHost);
			ToolbarAnimation.FloatUp(this.RightBottomToolbarHost);

			// Counting from here rather than from show-begin: the rise is the thing that
			// must not be disturbed, and this is the moment it starts - whether that was
			// the first frame or the fallback timer.
			this.StartAnalysisAfterEntry();
		}

		/// <summary>
		/// Starts the thumbnail analysis once the entry animation has had the render
		/// loop to itself.
		///
		/// These two used to be queued straight from show-begin, at Normal priority -
		/// which is above Render, so they ran before the rise's first frame and then
		/// kept the UI thread busy for the whole 0.5 seconds the rise needs. A WPF
		/// clock counts wall-clock time, so the animation was over before a single
		/// frame of it had been drawn: the toolbar simply appeared, which is the
		/// "there is no rise animation" report.
		///
		/// One tick of a Background timer, a beat after the rise starts. Background is
		/// below Render, so the tick cannot preempt the frames the animation is made
		/// of. Opening the thumbnail panel earlier than this still works: the panel's
		/// own open path starts the analysis if it has not started yet.
		/// </summary>
		private void StartAnalysisAfterEntry()
		{
			// A presentation that lasted less than the hold would otherwise leave this
			// tick behind to fire into the NEXT one and start the analysis at the wrong
			// moment - which is the moment the rise is supposed to have to itself.
			if (this._analysisDelayTimer != null)
			{
				this._analysisDelayTimer.Stop();
			}

			DispatcherTimer timer = new DispatcherTimer(DispatcherPriority.Background)
			{
				Interval = TimeSpan.FromMilliseconds(MainWindow.EntryRiseSettleMilliseconds)
			};
			this._analysisDelayTimer = timer;
			timer.Tick += delegate
			{
				timer.Stop();
				if (!PptOperator.Instance.IsShowing)
				{
					// The show ended inside the same beat; there is nothing to export
					// from any more and the next one starts its own pass.
					return;
				}
				// The thumbnails are fetched as soon as the show starts, not when the
				// panel is first opened. AnalyzePpt runs the export; the explicit
				// StartAnalyzeAllAsync here is what turns that into the full pass, and
				// putting it here rather than in ToggleThumbnailPanel is what makes the
				// panel show a complete list the first time it is opened.
				SyntacticSugar.CurrentDispatcherBeginInvoke(new Action(delegate
				{
					PptAnalyzer.CurrentShowPosition = this._currentPageIndex;
					PptAnalyzer.StartAnalyzeAllAsync(AnalyzeRate.Highest);
				}));

				SyntacticSugar.CurrentDispatcherBeginInvoke(new Action(this.AnalyzePpt));
			};
			timer.Start();
		}

		// ================= placement =================

		/// <summary>
		/// Clicking anywhere in the toolbars other than the menu's own button closes
		/// an open tool menu and returns to the tool that is active.
		///
		/// Hooked on the four side-toolbar hosts plus the centre one, in the tunnel
		/// (Preview) phase, so it sees the press before anything marks it handled.
		///
		/// A press that STARTED on one of the three tool buttons is ignored, and that
		/// check is not an optimisation - it is the whole correctness of this handler.
		/// Tunneling runs parent to child, so these host handlers run BEFORE the tool
		/// button's own PreviewMouseLeftButtonDown, not after as the obvious reading
		/// of "handled" suggests. When the dismissal ran first, the third press on a
		/// tool button closed its menu and then the button's own handler immediately
		/// reopened it, with the stale Closed event from the first half landing
		/// afterwards and resetting the open flag. Net effect: the pen colour menu
		/// would not stay open and the cycle stopped working.
		///
		/// So the button decides its own menu, and only a press somewhere else in the
		/// toolbar dismisses.
		/// </summary>
		private void HookToolMenuDismissal()
		{
			FrameworkElement[] hosts =
			{
				this.LeftToolbarHost,
				this.LeftBottomToolbarHost,
				this.CenterToolbarHost,
				this.RightToolbarHost,
				this.RightBottomToolbarHost,
				this.ToolPanelStack
			};

			foreach (FrameworkElement host in hosts)
			{
				// BUBBLE phase, not Preview.
				//
				// In the tunnel phase this handler runs before the button's own, so every
				// press on a tool button dismissed the panel and then the button's handler
				// opened it again in the same press - two opposing animations on one
				// property, and the collapse callback of the first never arrived.
				//
				// In the bubble phase it runs after, and a tool button press never reaches
				// it at all: BottomArea's own PreviewMouseLeftButtonDown sets Handled, and
				// a handled event does not bubble. The button therefore owns its menu
				// completely, and only a press elsewhere in the toolbars dismisses.
				host.MouseLeftButtonDown += delegate(object sender, MouseButtonEventArgs e)
				{
					if (ViewHelper.BottomArea == null)
					{
						return;
					}

					DependencyObject source = e.OriginalSource as DependencyObject;

					// Belt and braces: the eraser panel's slider does not mark anything
					// handled, and the press that starts a drag lands on it, so it must
					// not be read as a click away from the panel.
					if (BottomArea.IsToolButton(source) || ViewHelper.BottomArea.IsInsideToolPanel(source))
					{
						return;
					}

					ViewHelper.BottomArea.DismissToolMenus();
				};
			}
		}

		/// <summary>
		/// Lays the five hosting Grids out along the bottom of the presentation area:
		/// left, centre and right in a row, all on the same baseline.
		///
		/// This used to read five ToolbarPlacement entries from the settings, but the
		/// position controls were removed - the toolbars always sit at the bottom now,
		/// so this is the single place that says so. The two thumbnail panel hosts are
		/// placed here too: stretched over the row above the toolbars, hugging the left
		/// and right edges, and as wide as the side toolbar each one stands on, since
		/// they are stacked panels rather than part of the toolbar row.
		///
		/// The edge inset is now always zero, so nothing is scaled with the resolution
		/// any more; see GetEdgeInsetForWidth.
		/// </summary>
		public void ApplyToolbarPlacements()
		{
			double inset = this.GetEdgeInset();

			// the three toolbars: one row, hugging the bottom edge
			MainWindow.PinToBottom(this.LeftToolbarHost, HorizontalAlignment.Left, inset);
			MainWindow.PinToBottom(this.LeftBottomToolbarHost, HorizontalAlignment.Left, inset);
			MainWindow.PinToBottom(this.CenterToolbarHost, HorizontalAlignment.Center, inset);
			MainWindow.PinToBottom(this.RightToolbarHost, HorizontalAlignment.Right, inset);
			MainWindow.PinToBottom(this.RightBottomToolbarHost, HorizontalAlignment.Right, inset);

			// the thumbnail panels: standing on top of the toolbar row at the left and
			// right edges, growing upward into row 0.
			//
			// They are stretched over that row rather than centred in it, and the
			// panel inside is bottom-aligned and capped to this height - so the bottom
			// edge of a panel is always the top edge of the side toolbar, however many
			// slides there are. They used to sit at mid height instead, which left a
			// screen-wide gap under every open panel and meant a long deck was capped
			// at a fixed 470px no matter how much room there was.
			this.LeftThumbnailHost.HorizontalAlignment = HorizontalAlignment.Left;
			this.LeftThumbnailHost.VerticalAlignment = VerticalAlignment.Stretch;
			this.LeftThumbnailHost.Margin = new Thickness(inset, 0, 0, 0);
			this.RightThumbnailHost.HorizontalAlignment = HorizontalAlignment.Right;
			this.RightThumbnailHost.VerticalAlignment = VerticalAlignment.Stretch;
			this.RightThumbnailHost.Margin = new Thickness(0, 0, inset, 0);

			// ...and exactly as wide as that side toolbar.
			//
			// Measured, not written down. The toolbar is sized by its own content,
			// which changes with the resolution tier (icon size, font size, number of
			// buttons), so a constant here would be wrong for two tiers out of three -
			// and the mismatch is not cosmetic: an entry wider than the width the
			// ListBox template hands its ScrollViewer gets shaved at the viewport
			// edge, which is how the selection ring ended up drawn on three sides
			// instead of four.
			//
			// Office 2016 and newer use the horizontal side toolbars and older versions
			// the vertical ones. The pair that is not in use is an empty Collapsed Grid
			// and measures 0, so all four are tried in turn and NaN - which means Auto
			// - is the answer when none of them has been laid out yet. The left and
			// right toolbars are mirror images of each other, so one measurement is
			// good for both panels.
			double panelWidth = MainWindow.ToolbarWidth(
				this.LeftBottomToolbarHost,
				this.LeftToolbarHost,
				this.RightBottomToolbarHost,
				this.RightToolbarHost);

			// ...but never narrower than a legible entry needs.
			//
			// The vertical side toolbars of Office 2013 and older are 56px wide, less
			// than the 90px floor on an entry. Matching one exactly would put an entry
			// wider than the panel into it and have the ScrollViewer cut the entry - and
			// the selection ring with it - in half. Every toolbar from Office 2016 on is
			// 176px or more, so this floor never applies there.
			//
			// The room an entry needs is what one entry costs plus the list's own
			// padding, the scrollbar, and the pixel left against rounding - see
			// MinimumPanelWidth, which takes off the same things that
			// MeasureContentWidth does when it works the entry size out.
			double minimumPanelWidth = SlideThumbnailPanel.MinimumPanelWidth(MainWindow.MinimumSlideThumbnailWidth);
			if (!double.IsNaN(panelWidth) && panelWidth < minimumPanelWidth)
			{
				panelWidth = minimumPanelWidth;
			}

			MainWindow.SetHostWidth(this.LeftThumbnailHost, panelWidth);
			MainWindow.SetHostWidth(this.RightThumbnailHost, panelWidth);

			// The entry size is NOT settled here. It comes from what the panel's own
			// ScrollViewer leaves over, which is known only once the panel has been laid
			// out, and the panel raises ContentWidthChanged when that number moves. See
			// ThumbnailPanel_ContentWidthChanged.
		}

		/// <summary>
		/// The first of these hosts that has really been measured, or NaN.
		/// </summary>
		private static double ToolbarWidth(params FrameworkElement[] hosts)
		{
			foreach (FrameworkElement host in hosts)
			{
				double width = host.ActualWidth;
				if (!double.IsNaN(width) && width > 0.0)
				{
					return width;
				}
			}

			return double.NaN;
		}

		/// <summary>
		/// Writes a host width only when it really changed.
		///
		/// This runs from the size changed handler of the toolbar next to it, and an
		/// assignment of the same number still invalidates the measure pass - so a
		/// plain setter would keep re-measuring for a width that never moved.
		/// </summary>
		private static void SetHostWidth(FrameworkElement host, double width)
		{
			double current = host.Width;
			bool same = double.IsNaN(current)
				? double.IsNaN(width)
				: (!double.IsNaN(width) && Math.Abs(current - width) < 0.01);
			if (!same)
			{
				host.Width = width;
			}
		}

		private static void PinToBottom(FrameworkElement host, HorizontalAlignment horizontal, double inset)
		{
			host.HorizontalAlignment = horizontal;
			host.VerticalAlignment = VerticalAlignment.Bottom;
			host.Margin = new Thickness(0, 0, 0, inset);
		}

		/// <summary>Edge inset, scaled with the resolution.</summary>
		public static double GetEdgeInsetForWidth(double width)
		{
			// Zero, on purpose.
			//
			// This used to scale with the resolution - 6 / 12 / 36 - so the toolbars
			// would not sit flush against the bezel on a large display. The requirement
			// is the opposite: every control sits on the bottom edge with no gap at all,
			// so the row reads as part of the screen rather than floating above it.
			//
			// The parameter is kept so callers and the resolution tiers do not have to
			// change, and so a future "gap by tier" behaviour is a one-line edit here
			// rather than a change at every call site.
			return 0.0;
		}

		private double GetEdgeInset()
		{
			return MainWindow.GetEdgeInsetForWidth(base.Width);
		}

		// ================= show / hide =================

		/// <summary>All three toolbars float up (used when a full-screen panel closes).</summary>
		public void ShowAllPanels()
		{
			ToolbarAnimation.FloatUp(this.LeftToolbarHost);
			ToolbarAnimation.FloatUp(this.LeftBottomToolbarHost);
			ToolbarAnimation.FloatUp(this.CenterToolbarHost);
			ToolbarAnimation.FloatUp(this.RightToolbarHost);
			ToolbarAnimation.FloatUp(this.RightBottomToolbarHost);
			this.PlayShowLeftCenterButtonGridStoryboard();
			this.PlayShowRightCenterButtonGridStoryboard();
		}

		/// <summary>
		/// After a page change or a tool switch: fold the open thumbnail panels and
		/// dim the side toolbars, leaving only the center toolbar on screen.
		/// </summary>
		public void HideAllPanels()
		{
			this.HideThumbnailPanel(this.LeftThumbnailPanel);
			this.HideThumbnailPanel(this.RightThumbnailPanel);
			this.PlayHiddenLeftCenterButtonGridStoryboard();
			this.PlayHiddenRightCenterButtonGridStoryboard();
		}

		/// <summary>
		/// The thumbnail button's handler, shared by all four toolbars.
		///
		/// All four (left / right, vertical / horizontal) used to carry their own copy
		/// of this, and every copy began by refusing to do anything:
		///
		///     if (count <= 0 || count != PptOperator.Instance.SlidesCount) { log; return; }
		///
		/// The panel was therefore only ever openable in the one state where its
		/// contents already matched a second COM read. Combined with the missing
		/// ItemsSource binding (which pinned Items.Count at 0) that made the button
		/// do nothing at all, with nothing on screen to show that anything had been
		/// pressed.
		///
		/// So the panel is now always opened. A count that disagrees with SlidesCount
		/// is worth knowing about, but it is a log line, not a reason to hide the one
		/// control the user just asked for - and the count often disagrees legitimately,
		/// while the analysis is still running and fills the list in afterwards.
		/// </summary>
		public void ToggleThumbnailPanelFromToolbar(SlideThumbnailPanel panel)
		{
			if (panel == null)
			{
				return;
			}

			int count = panel.SlideListBox.Items.Count;
			int expected = PptOperator.Instance.SlidesCount;
			if (count != expected)
			{
				Log.WriteDebug(string.Format(
					"thumbnail panel opening with {0} items for {1} slides (analysis still running?)",
					count, expected));
			}

			this.ToggleThumbnailPanel(panel);
		}

		/// <summary>
		/// Thumbnail button: opens this panel, closes the other one and dims the side
		/// toolbars; pressing it again closes it.
		/// </summary>
		public void ToggleThumbnailPanel(SlideThumbnailPanel panel)
		{
			if (panel == null)
			{
				return;
			}

			// The state to test is the HOST's visibility, not the panel's. The panel
			// UserControl is never hidden - it stays Visible inside its host for its
			// whole life - and only the host is ever collapsed, by SlideOut when the
			// panel closes.
			// Reading panel.Visibility here meant it was always Visible, so every
			// press took the "close it" branch, which then found the host already
			// collapsed and returned: the panel could never be opened at all.
			FrameworkElement host = panel == this.LeftThumbnailPanel ? this.LeftThumbnailHost : this.RightThumbnailHost;
			if (host.Visibility == Visibility.Visible)
			{
				this.HideThumbnailPanel(panel);
				return;
			}

			// Opening one panel closes the other and dims the side toolbars so that
			// nothing overlaps the thumbnails.
			this.HideThumbnailPanel(panel == this.LeftThumbnailPanel ? this.RightThumbnailPanel : this.LeftThumbnailPanel);
			this.PlayHiddenLeftCenterButtonGridStoryboard();
			this.PlayHiddenRightCenterButtonGridStoryboard();

			// Slide in from the side the panel stands on, not up from below: the
			// panels are tall and at the screen edges, so a rise reads as a separate
			// box dropping in rather than as the toolbars opening.
			ToolbarAnimation.SlideIn(host, host == this.LeftThumbnailHost);
			this.SetThumbnailPanelShown(panel, true);
			SyntacticSugar.CurrentDispatcherBeginInvoke(DispatcherPriority.Background, new Action(delegate
			{
				this.HighMaskButton.Visibility = Visibility.Visible;
			}));
			WindowsApi.SetForegroundWindow(PptOperator.Instance.SlideShowWindowIntPtr);
			PptAnalyzer.CurrentShowPosition = this._currentPageIndex;

			// The analysis is already running from the start of the show (see
			// _pptOperatorShowBegin); this call is a no-op if it is, and covers the case
			// where the panel is opened before it has started.
			PptAnalyzer.StartAnalyzeAllAsync(AnalyzeRate.Highest);
		}

		private void HideThumbnailPanel(SlideThumbnailPanel panel)
		{
			if (panel == null)
			{
				return;
			}
			FrameworkElement host = panel == this.LeftThumbnailPanel ? this.LeftThumbnailHost : this.RightThumbnailHost;
			if (host.Visibility != Visibility.Visible)
			{
				return;
			}
			this.SetThumbnailPanelShown(panel, false);
			ToolbarAnimation.SlideOut(host, host == this.LeftThumbnailHost, delegate
			{
				host.Visibility = Visibility.Collapsed;
			});
		}

		private void SetThumbnailPanelShown(SlideThumbnailPanel panel, bool shown)
		{
			if (panel == this.LeftThumbnailPanel)
			{
				this.IsLeftSlideManagerViewShowed = shown;
			}
			else if (panel == this.RightThumbnailPanel)
			{
				this.IsRightSlideManagerViewShowed = shown;
			}
		}

		// ================= folding the side toolbars with the centre bar =================

		/// <summary>
		/// Folds or unfolds the four side toolbars together with the centre bar.
		///
		/// The centre bar keeps a key of its own when it folds (half a bar, press to
		/// bring everything back); the side toolbars have none. They simply sink,
		/// because what they hold - the page control and the thumbnail buttons - is
		/// also in the thumbnail panel, and one key that puts the whole row back is
		/// easier to find than five.
		///
		/// Called by BottomArea, which owns that key. It passes true in the same frame
		/// as the bar's own sink, and false only once the half bar has finished
		/// sinking: the full bar rises on that tick and these four rise with it, or the
		/// two ends of the row would arrive ahead of the middle they belong to.
		/// </summary>
		internal void SetSideToolbarsFolded(bool folded)
		{
			if (this._isSideToolbarsFolded == folded)
			{
				return;
			}
			this._isSideToolbarsFolded = folded;
			Log.WriteDebug(folded ? "side toolbars: folding" : "side toolbars: unfolding");

			FrameworkElement[] sideHosts =
			{
				this.LeftToolbarHost,
				this.LeftBottomToolbarHost,
				this.RightToolbarHost,
				this.RightBottomToolbarHost
			};

			if (folded)
			{
				// A thumbnail panel is opened from the side toolbars, so it goes down
				// with them: left open it would be a panel with no controller and no
				// way back to the row it belongs to.
				this.HideThumbnailPanel(this.LeftThumbnailPanel);
				this.HideThumbnailPanel(this.RightThumbnailPanel);

				foreach (FrameworkElement host in sideHosts)
				{
					ToolbarAnimation.FloatDown(host, null);
				}
				return;
			}

			foreach (FrameworkElement host in sideHosts)
			{
				ToolbarAnimation.FloatUp(host);
			}
		}

		// ================= idle dimming of the side toolbars =================

		private void MainWindow_LocationChanged(object sender, EventArgs e)
		{
			this.ResetPreWindowRect();
		}

		private void MainWindow_SizeChanged(object sender, SizeChangedEventArgs e)
		{
			this.AdjustFontSizeAndStyle();
			this.ApplyToolbarPlacements();
		}

		public void _pptOperatorShowBegin(object sender, EventArgs e)
		{
			SyntacticSugar.TryCatchAction(delegate
			{
				WindowStateHelper.UpdateWindowRect(this);
				SyntacticSugar.CurrentDispatcherInvoke(delegate
				{
					DispatcherTimer releaseTimer = this._releaseTimer;
					if (releaseTimer != null)
					{
						releaseTimer.Stop();
					}

					// The last show's exit sink must not fade this one's overlay out
					// from underneath it. It is only 300ms, but a presentation can
					// begin inside that - and the fade that would follow would take
					// the whole window to zero while this show is starting.
					DispatcherTimer exitSinkTimer = this._exitSinkTimer;
					if (exitSinkTimer != null)
					{
						exitSinkTimer.Stop();
					}
					this.EnsureToolbarsAttached();
					base.Show();
					base.Activate();
					this.ApplyToolbarPlacements();
					this.AdjustFontSizeAndStyle();
					BottomArea bottomArea = ViewHelper.BottomArea;
					if (bottomArea != null)
					{
						bottomArea.InitializeBottomArea();
					}
					this._currentPageIndex = PptOperator.Instance.CurrentShowPosition;
					this.SetPageIndex(this._currentPageIndex);

					// The hosts go up INVISIBLE, not at full opacity: the entry
					// animation is what reveals them, and showing them here would put
					// the whole toolbar on screen for a frame before the rise dropped it
					// back down again. SetToolbarsOpacity(1) used to live here.
					//
					// A fold from the last presentation is cleared in the same breath.
					// The side toolbars come back with the others (their hosts are
					// prepared above), and the centre bar is put back to its full width
					// by InitializeBottomArea a few lines up.
					this._isSideToolbarsFolded = false;
					this.SetEntryHostsPrepared(true);

					// And the thumbnail hosts, for a show that began while the last
					// one's exit sink was still carrying them down: collapsed, and
					// with the offset the sink left cleared, or the sideways slide
					// that opens them next would run underneath it.
					this.SettleExitedThumbnailHosts();

					this.IsExiting = false;
					base.Opacity = 1.0;
					Log.WriteDebug(string.Format(
						"show begin: windowLoaded={0} toolbarsAttached={1} size={2}x{3}",
						this.IsWindowLoaded, this._toolbarsAttached,
						base.Width, base.Height));
				});

				// Queued at Background - below Render - so it runs after the window has
				// settled and cannot push a frame of the rise out of the way. What it
				// does is arm the rise: the toolbars go on screen at zero opacity and
				// FloatUp starts on the first frame that actually draws them.
				base.Dispatcher.BeginInvoke(
					DispatcherPriority.Background,
					new Action(this.PlayEntryAnimation));

				// And the analysis waits behind the rise - see StartAnalysisAfterEntry,
				// which starts counting from the moment the rise actually starts.
			}, null);
		}

		private void _pptOperatorShowEnd(object sender, EventArgs e)
		{
			this.ShowEnd();
		}

		private void PptOperatorNextSlide(object sender, EventArgs e)
		{
			SyntacticSugar.CurrentDispatcherBeginInvoke(delegate
			{
				if (this.IsTransparent)
				{
					return;
				}
				this.UpdatePageIndex();
			});
		}

		private void HighMaskButton_Click(object sender, MouseButtonEventArgs e)
		{
			SyntacticSugar.TryCatchAction(delegate
			{
				this.HighMaskButton.Visibility = Visibility.Collapsed;
				this.HideThumbnailPanel(this.LeftThumbnailPanel);
				this.HideThumbnailPanel(this.RightThumbnailPanel);
				this.IsLeftSlideManagerViewShowed = false;
				this.IsRightSlideManagerViewShowed = false;
				WindowsApi.SetForegroundWindow(PptOperator.Instance.SlideShowWindowIntPtr);
			}, null);
		}

		private void ShowLeftCenterButtonGrid_OnCompleted(object sender, EventArgs e)
		{
			this.IsLeftSlideManagerViewShowed = false;
			SyntacticSugar.CurrentDispatcherBeginInvoke(DispatcherPriority.Background, new Action(this.UpdatePageIndex));
			WindowsApi.SetForegroundWindow(PptOperator.Instance.SlideShowWindowIntPtr);
		}

		private void HiddenLeftCenterButtonGrid_OnCompleted(object sender, EventArgs e)
		{
			WindowsApi.SetForegroundWindow(PptOperator.Instance.SlideShowWindowIntPtr);
		}

		private void ShowRightCenterButtonGrid_OnCompleted(object sender, EventArgs e)
		{
			this.IsRightSlideManagerViewShowed = false;
			SyntacticSugar.CurrentDispatcherBeginInvoke(DispatcherPriority.Background, new Action(this.UpdatePageIndex));
			WindowsApi.SetForegroundWindow(PptOperator.Instance.SlideShowWindowIntPtr);
		}

		private void HiddenRightCenterButtonGrid_OnCompleted(object sender, EventArgs e)
		{
			WindowsApi.SetForegroundWindow(PptOperator.Instance.SlideShowWindowIntPtr);
		}

		/// <summary>
		/// The dim/undim stories are skipped while a fold is on. They work on the side
		/// toolbars' OPACITY, and a folded row is not a faded row: it sits below the
		/// bottom edge of the window (see SetSideToolbarsFolded and ToolbarAnimation).
		/// Starting them here would put an opacity clock on four hosts nobody can see,
		/// and the undim story's own Completed handler takes the foreground back from
		/// PowerPoint - both things only make sense while the row is on screen. One
		/// thing at a time owns the side toolbars.
		/// </summary>
		public void PlayHiddenLeftCenterButtonGridStoryboard()
		{
			if (this._isSideToolbarsFolded)
			{
				return;
			}
			Storyboard storyboard = base.Resources["HiddenLeftCenterButtonGridStoryBoard"] as Storyboard;
			if (storyboard == null)
			{
				return;
			}
			storyboard.Begin();
		}

		public void PlayHiddenRightCenterButtonGridStoryboard()
		{
			if (this._isSideToolbarsFolded)
			{
				return;
			}
			Storyboard storyboard = base.Resources["HiddenRightCenterButtonGridStoryBoard"] as Storyboard;
			if (storyboard == null)
			{
				return;
			}
			storyboard.Begin();
		}

		private void PlayShowLeftCenterButtonGridStoryboard()
		{
			if (this._isSideToolbarsFolded)
			{
				return;
			}
			Storyboard storyboard = base.Resources["ShowLeftCenterButtonGridStoryBoard"] as Storyboard;
			if (storyboard == null)
			{
				return;
			}
			storyboard.Begin();
		}

		private void PlayShowRightCenterButtonGridStoryboard()
		{
			if (this._isSideToolbarsFolded)
			{
				return;
			}
			Storyboard storyboard = base.Resources["ShowRightCenterButtonGridStoryBoard"] as Storyboard;
			if (storyboard == null)
			{
				return;
			}
			storyboard.Begin();
		}

		// ================= page number / thumbnail sync =================

		public void UpdatePageIndex()
		{
			if (this._isAnalyzing)
			{
				return;
			}
			SyntacticSugar.TryCatchAction(delegate
			{
				this._currentPageIndex = PptOperator.Instance.CurrentShowPosition;
				if (this._currentPageIndex > PptOperator.Instance.SlidesCount)
				{
					return;
				}
				if (this._currentPageIndex <= 0)
				{
					this._currentPageIndex = 1;
				}
				this.SetPageIndex(this._currentPageIndex);
				this.UpdateSlideListBoxIndex(this._currentPageIndex);
			}, null);
		}

		private void UpdateButtonState()
		{
			PptOperator.Instance.UpdatePointerColor();
			// No argument: BottomArea reads the pointer type itself, so that a refused
			// COM call can be told apart from "the pointer is selected". Passing
			// PptOperator.PointerType in flattened both into None.
			ViewHelper.BottomArea.UpdateButtonState();
		}

		private void UpdateWindowOpacity()
		{
			if (this._isAnalyzing)
			{
				return;
			}
			bool pptShowingState = WindowStateHelper.GetPptShowingState(this);
			if (this.IsTransparent && pptShowingState)
			{
				this.UpdatePageIndex();
			}

			// This is window-level opacity: one false reading here takes the whole
			// overlay - toolbars, half bar and all - off the screen for as long as it
			// lasts, which is worth a line in the log whenever it happens.
			double target = (double)(pptShowingState ? 1 : 0);
			if (Math.Abs(base.Opacity - target) > 0.001)
			{
				Log.WriteDebug(string.Format(
					"overlay opacity {0:0.##} -> {1:0.##} (pptOnScreen={2}, isShowing={3})",
					base.Opacity, target, pptShowingState, PptOperator.Instance.IsShowing));
			}
			base.Opacity = target;
		}

		private void UpdateSlideListBoxIndex(int index)
		{
			this.CancelSubscribeSelectionChangedEvent();
			int target = index - 1;
			if (this.LeftThumbnailPanel != null)
			{
				this.SelectSlideInPanel(this.LeftThumbnailPanel, target);
			}
			if (this.RightThumbnailPanel != null)
			{
				this.SelectSlideInPanel(this.RightThumbnailPanel, target);
			}
			this.SubscribeSelectionChangedEvent();
		}

		private void SelectSlideInPanel(SlideThumbnailPanel panel, int index)
		{
			if (panel == null || panel.SlideListBox == null)
			{
				return;
			}
			if (panel.SlideListBox.SelectedIndex == index)
			{
				return;
			}
			panel.SlideListBox.SelectedIndex = index;
			panel.SlideListBox.ScrollIntoView(panel.SlideListBox.SelectedItem);
		}

		private void UpdateGestureAndMouseLineTrackState()
		{
			if (this.IsTransparent || this.IsLeftSlideManagerViewShowed || this.IsRightSlideManagerViewShowed || ViewHelper.BottomArea.IsNeedStopMouseLineTrack)
			{
				MouseHelper.Instance.Stop();
				return;
			}
			MouseHelper.Instance.Start();
		}

		public void HiddenWindow()
		{
			base.Width = 1.0;
			base.Height = 1.0;
			WindowStateHelper.PreWindowRect.Bottom = 1;
			DispatcherTask.Start(100.0, new Action(base.Hide));
		}

		private void AnalyzePpt()
		{
			if (this._isAnalyzing)
			{
				return;
			}
			this._isAnalyzing = true;
			SyntacticSugar.TryCatchAction(delegate
			{
				string pptFilePath = PptOperator.Instance.PptFilePath;
				List<string> list;
				bool flag = PptAnalyzer.FindAlreadyAnalyzedFiles(pptFilePath, out list);
				if ((this.SlideItems != null && pptFilePath == this._prePptFilePath && flag) || PptOperator.Instance.ShowRangeType == PpSlideShowRangeType.ppShowNamedSlideShow)
				{
					this.DoAfterAnalyzePpt();
					return;
				}
				if (flag)
				{
					this.SetListBoxItem(list);
					this.DoAfterAnalyzePpt();
				}
				else
				{
					// The loading screen is gone; only the initialisation and the
					// scheduling of the follow-up analysis are left.
					this.InitListBoxItem();
					SyntacticSugar.CurrentDispatcherBeginInvoke(new Action(this.DoAfterAnalyzePpt));
				}
				this._prePptFilePath = pptFilePath;
			}, null);
		}

		private void ShowEnd()
		{
			SyntacticSugar.TryCatchAction(delegate
			{
				SyntacticSugar.CurrentDispatcherBeginInvoke(delegate
				{
					// Whatever the entry animation was still waiting for is over, because
					// the show is. Leaving it armed would let a stray frame handler rise
					// the toolbars after the window has gone invisible - and the next
					// presentation would arm a second one next to it.
					this.StopEntryRendering();
					this.StopEntryFrameTimer();

					if (!ConfigHelper.Instance.IsNeedStart() && this.IsTransparent)
					{
						return;
					}

					// Everything on screen leaves together, down through the bottom
					// edge - see SinkEverythingOut. The window's own opacity reaches
					// zero only once that travel is over, on the timer SinkEverythingOut
					// starts, so the fade is no longer what ends the overlay: it is what
					// clears it away after the sink has carried it out of sight.
					this.SinkEverythingOut();

					this.UpdateTimer.Stop();
					MouseHelper.Instance.Stop();
					ViewHelper.BottomArea.ResetButtonState();
					this.IsLeftSlideManagerViewShowed = false;
					this.IsRightSlideManagerViewShowed = false;
					PptAnalyzer.StopAnalyzeAllAsyn();
					this._releaseTimer.Stop();
					this._releaseTimer.Start();
					this.IsExiting = false;
				});
			}, null);
		}

		/// <summary>
		/// How long the exit sink runs before the window itself fades out.
		/// Matches ToolbarAnimation's exit, as every other sink timer here does.
		/// </summary>
		private const int ExitSinkMilliseconds = 300;

		/// <summary>
		/// Fades the window out and settles what the exit sink left behind, or null
		/// while no show is ending.
		///
		/// A timer rather than an animation's Completed, for the reason
		/// <see cref="BottomArea"/> gives for the panel hides: WPF does not raise
		/// Completed when a later animation on the same property replaces this one, and
		/// an overlay left standing after the show has ended is the one outcome that
		/// must not depend on an event arriving. Stopped from show-begin, so a
		/// presentation starting inside the sink cannot have its overlay faded out from
		/// underneath it.
		/// </summary>
		private DispatcherTimer _exitSinkTimer;

		/// <summary>
		/// Takes the whole overlay down through the bottom edge of the window, the way
		/// a fold does, and starts the clock that fades the window out after it.
		///
		/// Show-end used to set the window's opacity to zero outright: the show ended
		/// and the toolbar was simply gone, with no movement to say where it had gone.
		/// This is the exit counterpart of <see cref="RaiseToolbars"/> - the five hosts
		/// sink, an open thumbnail panel sinks WITH them rather than sliding out
		/// sideways, and the tool panels sink through the bar they stand on. One
		/// movement, one duration, everything leaving on the same frame.
		///
		/// The four side hosts are skipped while they are already folded. Begin puts an
		/// element back in its slot before measuring it, so sinking one that is sitting
		/// below the edge would flash it on screen for a frame on its way down a second
		/// time - they are already where this is sending them.
		/// </summary>
		private void SinkEverythingOut()
		{
			// A second ShowEnd while the first sink is still running - the COM event
			// and the "ended unexpectedly" path can both arrive for one show - would
			// restart every travel from the slot, and Begin puts an element back in
			// its slot before measuring it: the row would flash on screen for a frame
			// on its way down a second time. On the way out already means done here.
			if (this._exitSinkTimer != null && this._exitSinkTimer.IsEnabled)
			{
				return;
			}

			Log.WriteDebug("show end: sinking the overlay out");

			// The mask an open thumbnail panel puts over the slide covers the whole
			// window and has nowhere to sink to, so it would simply sit on top of the
			// exit - and swallow the clicks meant for the panel while it plays.
			this.HighMaskButton.Visibility = Visibility.Collapsed;

			if (!this._isSideToolbarsFolded)
			{
				ToolbarAnimation.FloatDown(this.LeftToolbarHost, null);
				ToolbarAnimation.FloatDown(this.LeftBottomToolbarHost, null);
				ToolbarAnimation.FloatDown(this.RightToolbarHost, null);
				ToolbarAnimation.FloatDown(this.RightBottomToolbarHost, null);
			}

			// As a host rather than as a bar: the half bar lives inside this host, so
			// one travel carries whichever of the two halves is on screen. Folded or
			// not, the middle of the row leaves as the one piece it is showing.
			ToolbarAnimation.FloatDown(this.CenterToolbarHost, null);

			// With the row rather than on its own sideways slide-out: an exit is one
			// movement, and the panels stand at the screen edges the row is leaving
			// through anyway. The panel STATE clears here; the hosts themselves are
			// collapsed and their offsets cleared once the travel is over, in
			// StartExitSinkTimer - the next open is a sideways slide, which animates X
			// only and would never clear a vertical offset left under it.
			this.SinkThumbnailPanel(this.LeftThumbnailPanel);
			this.SinkThumbnailPanel(this.RightThumbnailPanel);

			// The colour and eraser panels stand ON the bar, so they have their own
			// travel - the same close animation a normal close plays, only started on
			// this frame with everything else.
			ViewHelper.BottomArea.SinkToolPanelsForExit();

			this.StartExitSinkTimer();
		}

		/// <summary>
		/// Puts a thumbnail panel's host on the exit sink, if that panel is open.
		///
		/// The same test <see cref="HideThumbnailPanel"/> makes - the HOST's visibility,
		/// because the panel itself never hides - and the same state clearing, but the
		/// travel is the row's own sink: on the way out these panels go where
		/// everything else goes, down, instead of sideways.
		/// </summary>
		private void SinkThumbnailPanel(SlideThumbnailPanel panel)
		{
			if (panel == null)
			{
				return;
			}

			FrameworkElement host = panel == this.LeftThumbnailPanel ? this.LeftThumbnailHost : this.RightThumbnailHost;
			if (host.Visibility != Visibility.Visible)
			{
				return;
			}

			this.SetThumbnailPanelShown(panel, false);
			ToolbarAnimation.FloatDown(host, null);
		}

		/// <summary>
		/// Arms the fade-out for when the exit sink has finished.
		/// </summary>
		private void StartExitSinkTimer()
		{
			if (this._exitSinkTimer == null)
			{
				this._exitSinkTimer = new DispatcherTimer
				{
					Interval = TimeSpan.FromMilliseconds(MainWindow.ExitSinkMilliseconds)
				};
				this._exitSinkTimer.Tick += delegate
				{
					this._exitSinkTimer.Stop();
					Log.WriteDebug("show end: sink over, overlay to zero opacity");
					base.Opacity = 0.0;
					this.SettleExitedThumbnailHosts();
				};
			}

			this._exitSinkTimer.Stop();
			this._exitSinkTimer.Start();
		}

		/// <summary>
		/// Collapses the two thumbnail hosts and clears the offsets the exit sink left
		/// in them.
		///
		/// The five toolbar hosts need none of this: show-begin prepares them and the
		/// rise cancels whatever a sink left behind. The thumbnail hosts are opened by
		/// a sideways slide, which animates X only, so a vertical offset left under it
		/// would keep the panel below the edge where nobody could see it. The same
		/// settle runs from show-begin, for a presentation that starts inside the sink.
		/// </summary>
		private void SettleExitedThumbnailHosts()
		{
			MainWindow.SettleExitedThumbnailHost(this.LeftThumbnailHost);
			MainWindow.SettleExitedThumbnailHost(this.RightThumbnailHost);
		}

		private static void SettleExitedThumbnailHost(FrameworkElement host)
		{
			if (host == null)
			{
				return;
			}

			host.Visibility = Visibility.Collapsed;

			// Opacity too: a sideways slide-out that the sink interrupted leaves its
			// own fade clock running over this, and the resting value is what the next
			// open has to find underneath it.
			host.BeginAnimation(UIElement.OpacityProperty, null);
			host.Opacity = 1.0;

			TranslateTransform transform = host.RenderTransform as TranslateTransform;
			if (transform != null)
			{
				transform.BeginAnimation(TranslateTransform.XProperty, null);
				transform.BeginAnimation(TranslateTransform.YProperty, null);
				transform.X = 0.0;
				transform.Y = 0.0;
			}
		}

		private void ResetPreWindowRect()
		{
			WindowStateHelper.PreWindowRect = new WindowRect
			{
				Left = (int)base.Left,
				Top = (int)base.Top,
				Right = (int)(base.Left + base.Width),
				Bottom = (int)(base.Top + base.Height)
			};
		}

		/// <summary>
		/// Picks the thumbnail style tier for the current resolution.
		///
		/// Only the styles are dealt with here. The entry size used to be chosen by
		/// this method as well - one hardcoded width per tier - but it is measured
		/// from the side toolbar the panel matches in width, and it follows the panel:
		/// see ApplyToolbarPlacements and <see cref="UpdateSlideThumbnailSize"/>.
		/// </summary>
		private void AdjustFontSizeAndStyle()
		{
			string listStyle = "Style.ThumbnailList1440";
			if (base.Width > 3800)
			{
				listStyle = "Style.ThumbnailList3800";
			}
			else if (base.Width <= 1440)
			{
				listStyle = "Style.ThumbnailList0";
			}

			if (this.LeftThumbnailPanel != null)
			{
				this.LeftThumbnailPanel.ApplyStyle(listStyle);
			}
			if (this.RightThumbnailPanel != null)
			{
				this.RightThumbnailPanel.ApplyStyle(listStyle);
			}

			// Settle the entry size here as well, not only in ApplyToolbarPlacements.
			//
			// The tier styles differ in their entry HEIGHT as well as their font size,
			// and the panel's own width does not change when the tier does - so
			// switching tiers is one of the few things that can move the height without
			// moving the width. Leaving it to the placement would leave the entries a
			// frame behind the style that was just applied to them.
			this.UpdateSlideThumbnailSize();
		}

		/// <summary>
		/// Floor for the entry width. Below this a thumbnail is too small to recognise
		/// at all, and the number above it stops being legible.
		/// </summary>
		private const double MinimumSlideThumbnailWidth = 90.0;

		/// <summary>Ceiling for the entry width, so a 4K panel does not run off the screen.</summary>
		private const double MaximumSlideThumbnailWidth = 400.0;

		/// <summary>
		/// Size of one thumbnail entry: the width the panel reports as available
		/// (its own width, less the list's padding and the scrollbar), and a height
		/// that matches the showing deck's own aspect ratio.
		///
		/// Both are computed rather than written down.
		///
		/// The width is derived from the panel, which MainWindow has just made as
		/// wide as the side toolbar; a fixed number per tier would be wrong for any
		/// screen where that toolbar did not come out at exactly the width the number
		/// was written for, and an entry even slightly wider than the panel is cut off
		/// by the ScrollViewer - a shaved left or right edge on a selected entry is
		/// exactly what a ring on three sides looks like. Hence the floor to whole
		/// pixels: the entry also has SnapsToDevicePixels, and a fractional width
		/// turns into a shaved edge there as well.
		///
		/// The height has to be computed for the same reason from the other end: the
		/// entry templates used to carry a fixed Height and the exported images carry
		/// the deck's real ratio, so any fixed number is wrong for every deck that is
		/// not that ratio and Image Stretch="Fill" then squashes or stretches every
		/// thumbnail to fit it.
		/// </summary>
		private void UpdateSlideThumbnailSize()
		{
			// The panel reports the width an entry may use, and it has already taken off
			// its own padding and the scrollbar. Floor to whole pixels: an entry on a
			// fractional width is rounded by the ScrollViewer against the viewport, and
			// that is what shaves the last pixel off the right edge of the selection ring.
			//
			// Until a panel has reported one, the dependency properties keep their
			// defaults - which are also the fallbacks the entry templates declare - so
			// the first frame draws at the middle tier's size rather than at nothing.
			double width = double.NaN;
			if (this.LeftThumbnailPanel != null)
			{
				width = this.LeftThumbnailPanel.ContentWidth;
			}
			if ((double.IsNaN(width) || width <= 0.0) && this.RightThumbnailPanel != null)
			{
				width = this.RightThumbnailPanel.ContentWidth;
			}
			if (double.IsNaN(width) || width <= 0.0)
			{
				return;
			}

			width = Math.Floor(width);
			if (width < MainWindow.MinimumSlideThumbnailWidth)
			{
				width = MainWindow.MinimumSlideThumbnailWidth;
			}
			else if (width > MainWindow.MaximumSlideThumbnailWidth)
			{
				width = MainWindow.MaximumSlideThumbnailWidth;
			}

			double aspect = PptOperator.FallbackSlideAspect;
			PptOperator pptOperator = PptOperator.Instance;
			if (pptOperator != null)
			{
				// The one COM read this method does. It cannot throw: ComLate returns
				// a fallback rather than propagating, so a closed presentation or a
				// product that will not answer leaves 16:9 in place.
				double reported = pptOperator.SlideAspectRatio;
				if (!double.IsNaN(reported) && reported > 0.0)
				{
					aspect = reported;
				}
			}

			// Rounded so a fractional height cannot leave half a pixel of white under
			// the ring, and floored at 1 so a refused ratio cannot collapse an entry.
			double height = Math.Round(width / aspect);
			if (height < 1.0)
			{
				height = 1.0;
			}

			// Both are only written when they changed: this runs on every resize, and
			// each write invalidates the measure of every entry in both lists.
			if (Math.Abs(this.SlideThumbnailWidth - width) > 0.5)
			{
				this.SlideThumbnailWidth = width;
			}
			if (Math.Abs(this.SlideThumbnailHeight - height) > 0.5)
			{
				this.SlideThumbnailHeight = height;
			}
		}

		/// <summary>
		/// Width of one thumbnail image, bound by Style.ThumbnailItem* in
		/// Style_\ToolbarStyle.xaml. Follows the side toolbar the panel is as wide as,
		/// so it changes with the resolution tier on its own rather than because a
		/// tier table says so. The default is the middle tier, which is also the
		/// FallbackValue the entry templates carry.
		/// </summary>
		public double SlideThumbnailWidth
		{
			get
			{
				return (double)base.GetValue(MainWindow.SlideThumbnailWidthProperty);
			}
			set
			{
				base.SetValue(MainWindow.SlideThumbnailWidthProperty, value);
			}
		}

		/// <summary>
		/// Height of one thumbnail entry, bound by Style.ThumbnailItem* in
		/// Style_\ToolbarStyle.xaml. <see cref="SlideThumbnailWidth"/> divided by the
		/// showing deck's aspect ratio; the default is 16:9 at the middle tier's width.
		/// </summary>
		public double SlideThumbnailHeight
		{
			get
			{
				return (double)base.GetValue(MainWindow.SlideThumbnailHeightProperty);
			}
			set
			{
				base.SetValue(MainWindow.SlideThumbnailHeightProperty, value);
			}
		}

		private void DoAfterAnalyzePpt()
		{
			SyntacticSugar.TryCatchAction(delegate
			{
				this.UpdateTimer.Start();
				this._isAnalyzing = false;
				MouseHelper.Instance.Start();
				WindowsApi.SetForegroundWindow(PptOperator.Instance.SlideShowWindowIntPtr);
			}, null);
		}

		private void SubscribeSelectionChangedEvent()
		{
			if (this.LeftThumbnailPanel != null)
			{
				this.LeftThumbnailPanel.SlideListBox.SelectionChanged += this.SlideListBox_SelectionChanged;
			}
			if (this.RightThumbnailPanel != null)
			{
				this.RightThumbnailPanel.SlideListBox.SelectionChanged += this.SlideListBox_SelectionChanged;
			}
		}

		private void CancelSubscribeSelectionChangedEvent()
		{
			if (this.LeftThumbnailPanel != null)
			{
				this.LeftThumbnailPanel.SlideListBox.SelectionChanged -= this.SlideListBox_SelectionChanged;
			}
			if (this.RightThumbnailPanel != null)
			{
				this.RightThumbnailPanel.SlideListBox.SelectionChanged -= this.SlideListBox_SelectionChanged;
			}
		}

		/// <summary>Overall opacity of the three toolbars (1 = fully visible).</summary>
		private void SetToolbarsOpacity(double value)
		{
			this.LeftToolbarHost.Opacity = value;
			this.LeftBottomToolbarHost.Opacity = value;
			this.CenterToolbarHost.Opacity = value;
			this.RightToolbarHost.Opacity = value;
			this.RightBottomToolbarHost.Opacity = value;
		}

		private void SetPageIndex(int index)
		{
			if (index <= 0)
			{
				index = 1;
			}
			if (PptOperator.Instance.ShowRangeType == PpSlideShowRangeType.ppShowSlideRange)
			{
				index += PptOperator.Instance.StartingSlide - 1;
			}
			string text = ((PptOperator.Instance.ShowRangeType == PpSlideShowRangeType.ppShowNamedSlideShow) ? string.Format("{0}", index) : string.Format("{0}/{1}", index, PptOperator.Instance.SlidesCount));
			ViewHelper.LeftArea.LeftSlideIndexTextBlock.Text = text;
			ViewHelper.RightArea.RightSlideIndexTextBlock.Text = text;
			ViewHelper.LeftBottomArea.LeftSlideIndexTextBlock.Text = text;
			ViewHelper.RightBottomArea.RightSlideIndexTextBlock.Text = text;
		}

		private void InitListBoxItem()
		{
			string[] array = new string[PptOperator.Instance.SlidesCount];
			foreach (KeyValuePair<int, string> keyValuePair in PptAnalyzer.AnalyzedIndexAndThumbnailFilePathDictionary)
			{
				array[keyValuePair.Key - 1] = keyValuePair.Value;
			}
			this.SetListBoxItem(array.ToList<string>());
		}

		private void SetListBoxItem(List<string> files)
		{
			if (files == null || !files.Any<string>())
			{
				return;
			}
			this.CancelSubscribeSelectionChangedEvent();
			this.SlideItems = new ObservableCollection<SlideItem>();
			if (PptOperator.Instance.ShowRangeType == PpSlideShowRangeType.ppShowAll)
			{
				int num = 1;
				foreach (string text in files)
				{
					this.CreateSlideItem(num++, text);
				}
			}
			if (PptOperator.Instance.ShowRangeType == PpSlideShowRangeType.ppShowSlideRange)
			{
				for (int i = PptOperator.Instance.StartingSlide; i <= PptOperator.Instance.EndingSlide; i++)
				{
					this.CreateSlideItem(i, files[i - 1]);
				}
			}
			this.SubscribeSelectionChangedEvent();
		}

		public void UpdateListBoxItem(List<string> files)
		{
			if (this.SlideItems == null)
			{
				return;
			}
			for (int i = 0; i < this.SlideItems.Count; i++)
			{
				if (files[i] != null && !(this.SlideItems[i].IsPptThumbnail == "True"))
				{
					BitmapImage bitmapImage = new BitmapImage(new Uri(files[i]));
					this.SlideItems[i].BitmapImage = bitmapImage;
					this.SlideItems[i].IsPptThumbnail = "True";
				}
			}
		}

		private void CreateSlideItem(int index, string filePath)
		{
			bool flag = !string.IsNullOrEmpty(filePath);
			SlideItem slideItem = new SlideItem
			{
				Index = index,
				IsPptThumbnail = flag.ToString(),
				BitmapImage = (flag ? new BitmapImage(new Uri(filePath)) : null)
			};
			this.SlideItems.Add(slideItem);
		}

		private void UpdateTimerTick(object sender, EventArgs e)
		{
			if (this.IsExiting || !this.IsNeedUpdateTimer || !PptOperator.Instance.IsShowing)
			{
				return;
			}
			SyntacticSugar.TryCatchAction(delegate
			{
				if (!this.IsNeedUpdateTimer)
				{
					return;
				}
				this.UpdateButtonState();
				if (!this.IsNeedUpdateTimer)
				{
					return;
				}
				this.UpdateWindowOpacity();
				this.UpdateGestureAndMouseLineTrackState();
				WindowStateHelper.UpdateWindowRect(this);
				if (base.Opacity > 0.0)
				{
					WindowStateHelper.MakeAppAbovePpt(this);
				}
			}, null);
		}

		private void MainWindow_KeyDown(object sender, KeyEventArgs e)
		{
			// App.StartProcess() assigns PptOperator.Instance, so it is not null in the
			// normal flow; the guard keeps an unlucky key press off the COM objects.
			PptOperator pptOperator = PptOperator.Instance;
			if (pptOperator == null)
			{
				return;
			}
			if (e.Key == Key.Escape)
			{
				// Same route as the on-screen exit button and the tray menu, so the
				// two-stage whiteboard exit and the IsExiting flag stay consistent.
				BottomArea bottomArea = ViewHelper.BottomArea;
				if (bottomArea != null)
				{
					bottomArea.ExitApplication();
				}
				else
				{
					pptOperator.Exit();
				}
				e.Handled = true;
				return;
			}
			if (e.Key == Key.Prior || e.Key == Key.Left || e.Key == Key.Up)
			{
				pptOperator.PreviousPage();
			}
			if (e.Key == Key.Next || e.Key == Key.Right || e.Key == Key.Down)
			{
				pptOperator.NextPage();
			}
		}

		private void SlideListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
		{
			SyntacticSugar.TryCatchAction(delegate
			{
				ListBox listBox = sender as ListBox;
				if (listBox == null)
				{
					return;
				}
				int num = listBox.SelectedIndex + 1;
				if (num <= 0)
				{
					return;
				}
				this.SetPageIndex(num);
				PptOperator.Instance.GotoSlide(num);
			}, null);
		}

		/// <summary>
		/// A panel has settled on the width one entry may use.
		///
		/// Fired from the panel's own SizeChanged, i.e. from inside a layout pass, and
		/// it re-measures every entry in both lists - so it must not fire for a width
		/// that did not really change. The panel checks that before raising it; see
		/// SlideThumbnailPanel.UpdateContentWidth.
		/// </summary>
		private void ThumbnailPanel_ContentWidthChanged()
		{
			SyntacticSugar.TryCatchAction(delegate
			{
				this.UpdateSlideThumbnailSize();
			}, null);
		}

		/// <summary>
		/// A thumbnail or a page number was clicked.
		///
		/// Only the slide that is already showing reaches this: a click on any other
		/// entry changes the selection, and SelectionChanged does the rest. But a click
		/// on the current one changes nothing about the selection and so raises nothing,
		/// and "click the slide I am already on to bring it back" is a reasonable thing
		/// for a user to do after an accidental drag in PowerPoint - so it is answered
		/// here rather than dropped.
		/// </summary>
		private void ThumbnailPanel_SlideRequested(int slideNumber)
		{
			SyntacticSugar.TryCatchAction(delegate
			{
				if (slideNumber <= 0)
				{
					return;
				}
				this.SetPageIndex(slideNumber);
				PptOperator.Instance.GotoSlide(slideNumber);
			}, null);
		}

		private void ReleaseTimerTick(object sender, EventArgs e)
		{
			SyntacticSugar.TryCatchAction(delegate
			{
				this._releaseTimer.Stop();
				this.ReleaseThumbnailImages();
				this.CancelSubscribeSelectionChangedEvent();
				this.SlideItems = null;
				this.SubscribeSelectionChangedEvent();
				GarbageCollection.Collect();
			},
			// The window is hidden in the finally action, not in the try block.
			//
			// It used to be the second-to-last line of the try, and it never ran: the
			// line above it builds a 1x1 placeholder bitmap, writing it to disk first
			// and then loading it back. Either half can throw (the folder is wiped at
			// start-up by AppPath.ClearTempFolder, so the write races its recreation;
			// a share violation or a read-only path does the same), and one throw was
			// enough to skip everything below - including the hide. The overlay was
			// then left on screen at zero opacity after every presentation, which is
			// what "the window does not hide itself when the show ends" was.
			//
			// Hiding the window is the one thing that must not depend on the memory
			// cleanup succeeding, so it is the one thing done unconditionally.
			delegate
			{
				this.HiddenWindow();
			});
		}

		/// <summary>
		/// Replaces every thumbnail with a shared 1x1 image so the decoded bitmaps can
		/// be collected. Failures are swallowed: this is only an optimisation, and it
		/// must never be the reason the toolbar window stays on screen.
		/// </summary>
		private void ReleaseThumbnailImages()
		{
			try
			{
				if (this.SlideItems == null || !this.SlideItems.Any<SlideItem>())
				{
					return;
				}

				BitmapImage bitmapImage = new BitmapImage(new Uri(ReleaseMemoryImageHelper.GetImageFilePath()));
				foreach (SlideItem slideItem in this.SlideItems)
				{
					slideItem.BitmapImage = bitmapImage;
				}
			}
			catch (Exception ex)
			{
				Log.WriteDebug("ReleaseThumbnailImages: " + ex.GetType().Name + ": " + ex.Message);
			}
		}

		private void QrGrid_OnMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
		{
			this.QrImage.Visibility = Visibility.Collapsed;
			this.QrGrid.Visibility = Visibility.Collapsed;
		}

		public ObservableCollection<SlideItem> SlideItems
		{
			get
			{
				return (ObservableCollection<SlideItem>)base.GetValue(MainWindow.SlideItemsProperty);
			}
			set
			{
				base.SetValue(MainWindow.SlideItemsProperty, value);
			}
		}

		private Gesture _gesture;
		private bool _isAnalyzing;
		private int _currentPageIndex;
		private string _prePptFilePath;

		private static readonly int TimerSpan = 300;
		private static readonly int ReleaseTimerSpan = 6000;

		/// <summary>
		/// How long the entry animation waits for a rendered frame before starting
		/// anyway, so a window that never presents one cannot leave the toolbars at
		/// zero opacity. See MainWindow.StartEntryAnimationOnFirstFrame.
		/// </summary>
		private const int EntryAnimationFallbackMilliseconds = 1000;

		/// <summary>
		/// How long the thumbnail analysis waits after the rise starts, so the 0.5
		/// second rise owns the render loop. See StartAnalysisAfterEntry.
		/// </summary>
		private const int EntryRiseSettleMilliseconds = 700;

		private readonly DispatcherTimer _releaseTimer;

		public static readonly DependencyProperty SlideItemsProperty = DependencyProperty.Register("SlideItems", typeof(ObservableCollection<SlideItem>), typeof(MainWindow), new PropertyMetadata(null));

		/// <summary>
		/// Backing store for <see cref="SlideThumbnailWidth"/>. 190 is the middle
		/// tier, which is what the entries fell back to before anything was measured.
		/// </summary>
		public static readonly DependencyProperty SlideThumbnailWidthProperty =
			DependencyProperty.Register("SlideThumbnailWidth", typeof(double),
				typeof(MainWindow), new PropertyMetadata(190.0));

		/// <summary>
		/// Backing store for <see cref="SlideThumbnailHeight"/>. 107 is 190 at 16:9,
		/// rounded - again the middle tier.
		/// </summary>
		public static readonly DependencyProperty SlideThumbnailHeightProperty =
			DependencyProperty.Register("SlideThumbnailHeight", typeof(double),
				typeof(MainWindow), new PropertyMetadata(107.0));
	}
}
