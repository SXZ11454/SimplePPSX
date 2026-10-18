using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Threading;

namespace SimplePPSX
{
	/// <summary>
	/// Thumbnail panel: one Rectangle backdrop plus one transparent ListBox.
	///
	/// It sizes itself, and it decides what a click means. Everything else - the
	/// entry styles, the show / hide animation and where on the screen it stands -
	/// is MainWindow's job (ToolbarAnimation and MainWindow.ApplyToolbarPlacements),
	/// so this one control serves whichever panel is asked for.
	/// </summary>
	public partial class SlideThumbnailPanel : UserControl
	{
		/// <summary>
		/// Room left above the panel, so a full-height one does not run into the top
		/// edge of the screen.
		/// </summary>
		private const double TopGap = 6.0;

		/// <summary>
		/// Floor for the height cap, used only until the host has been measured.
		/// Without it the panel's MaxHeight would be zero for the first frame and the
		/// list would flash collapsed.
		/// </summary>
		private const double MinimumCap = 120.0;

		/// <summary>
		/// Stands in for the scrollbar's width when none has been measured yet.
		///
		/// Deliberately the fat end of what a scrollbar can be rather than the 12px
		/// this app's one is drawn at: this number is only ever subtracted when the
		/// real one is unknown, and subtracting too much costs an entry a few pixels
		/// of width while subtracting too little lets the scrollbar land on top of the
		/// entry - which is a selection ring with one of its four sides missing. The
		/// wrong side of that trade is the cheap one.
		/// </summary>
		private const double ScrollBarAllowance = 20.0;

		/// <summary>
		/// Room left between an entry and the edge of the list's viewport, beyond the
		/// room the scrollbar already takes.
		///
		/// It exists because both the entry and the viewport are rounded to device
		/// pixels when they are drawn: an entry that measures exactly as wide as the
		/// room it was given can still be drawn one pixel wider than it, and one
		/// shaved pixel on the selected entry is a ring with one of its four sides
		/// missing.
		/// </summary>
		private const double EntrySlack = 1.0;

		/// <summary>
		/// Which side of its image a page number stands on. True - the default,
		/// and what the left panel is given in MainWindow - keeps the number on
		/// the image's left; the right panel leaves it false and the entry
		/// template moves the number column to the image's right.
		///
		/// A dependency property because the entry templates read it up the
		/// tree (Style_\ToolbarStyle.xaml, Style.ThumbnailItem*).
		/// </summary>
		public bool IsLeftSide
		{
			get
			{
				return (bool)base.GetValue(SlideThumbnailPanel.IsLeftSideProperty);
			}
			set
			{
				base.SetValue(SlideThumbnailPanel.IsLeftSideProperty, value);
			}
		}

		public static readonly DependencyProperty IsLeftSideProperty =
			DependencyProperty.Register("IsLeftSide", typeof(bool),
				typeof(SlideThumbnailPanel), new PropertyMetadata(false));

		/// <summary>
		/// Set once a measurement has run at least once. Before that the first one is
		/// deferred to after the initial layout pass; from then on it runs inline.
		/// </summary>
		private bool _loaded;

		/// <summary>Set while a deferred measurement is already queued.</summary>
		private bool _measureQueued;

		/// <summary>
		/// The list's vertical scrollbar, once we have found it, so its width can be
		/// followed and re-read when it moves. Replaced whenever the list re-templates
		/// itself (a tier switch changes its Style, and with it its template), which is
		/// what the reference check in <see cref="ScrollBarWidth"/> is for.
		/// </summary>
		private ScrollBar _scrollBar;

		public SlideThumbnailPanel()
		{
			// Matches the documented meaning of ContentWidth: nothing has been
			// measured yet, as opposed to "measured as zero".
			this.ContentWidth = double.NaN;

			this.InitializeComponent();
			// The backdrop joins the same chrome group as the toolbars, so one opacity
			// setting covers the panel too (see ToolbarAppearance).
			ToolbarAppearance.Register(this.ChromeRectangle);
			// The panel counts as a toolbar surface for hover too: while the
			// pointer is over it the chrome uses the default opacity, and leaving
			// glides back to the idle ("not clicked") one.
			ToolbarAppearance.HookHover(this);

			// The click that jumps to a slide is handled here rather than left to the
			// ListBox. See the two handlers for why.
			this.SlideListBox.PreviewMouseLeftButtonDown += this.OnSlideListPreviewMouseLeftButtonDown;
			this.SlideListBox.PreviewMouseLeftButtonUp += this.OnSlideListPreviewMouseLeftButtonUp;
			this.RootGrid.PreviewMouseLeftButtonDown += this.OnPanelPreviewMouseLeftButtonDown;

			// The width an entry may use follows this panel's own size, so it has to be
			// re-measured whenever that changes. See ContentWidth.
			this.SizeChanged += delegate
			{
				this.UpdateContentWidth();
			};

			// The floating selection ring: this panel owns it, and it is this
			// panel's selection that moves it - a glide when the pick moves, one
			// unmoving pass for everything else. See UpdateSelectionRing.
			this.SlideListBox.SelectionChanged += this.OnSlideListSelectionChanged;
			this.IsVisibleChanged += this.OnThumbnailPanelIsVisibleChanged;

			// Entries move when the panel resizes (its height follows the host's),
			// so the ring has to move with them - in the same frame, never on a
			// glide that would lag the entry it stands on.
			this.SizeChanged += delegate (object resizeSender, SizeChangedEventArgs resizeArgs)
			{
				this.UpdateSelectionRing(false);
			};

			this.Loaded += this.OnPanelLoaded;
		}

		/// <summary>
		/// Width one entry may use inside this panel, or NaN before the panel has been
		/// measured.
		///
		/// It is what the list's ScrollViewer leaves over, minus the room the scrollbar
		/// needs - see <see cref="MeasureContentWidth"/> for the whole sum, which is
		/// where the part that matters is explained.
		///
		/// The owner reads this to size the entries (MainWindow.UpdateSlideThumbnailSize)
		/// and listens for ContentWidthChanged to know when to do it again.
		/// </summary>
		public double ContentWidth { get; private set; }

		/// <summary>
		/// Raised when <see cref="ContentWidth"/> has changed.
		///
		/// This is the only way the owner finds out: the panel's width is set by the
		/// owner itself, so nothing it already listens to fires when the answer
		/// changes - the toolbar hosts keep their size, and the window does too.
		/// </summary>
		public event Action ContentWidthChanged;

		/// <summary>
		/// Raised when the user clicks a thumbnail or its page number, with the slide
		/// number to go to (1-based).
		///
		/// MainWindow owns the actual move: it already had one path for it
		/// (SlideListBox_SelectionChanged, off SelectionChanged), and a separate event
		/// is needed because a click on the slide that is already showing changes
		/// nothing about the selection and so would raise nothing at all.
		/// </summary>
		public event Action<int> SlideRequested;

		/// <summary>
		/// Raised on a press on the panel's own backing rather than on the list, so the
		/// owner can close it.
		///
		/// It used not to need raising: the full-screen mask covered the panel too, so a
		/// press anywhere at all closed it. The panel now sits above that mask so its
		/// entries can be clicked, which leaves the strip of backing around the list
		/// with nothing above it - a press there would otherwise do nothing at all.
		/// </summary>
		public event Action DismissRequested;

		/// <summary>
		/// Switches the resolution tier (0 / 1440 / 3800) by handing the list the
		/// matching style.
		///
		/// The width is not set here any more. The panel used to work its own width
		/// out of a per-tier table of thumbnail widths, which meant a second copy of
		/// numbers that also appear in Style_\ToolbarStyle.xaml; it is now as wide as
		/// the host MainWindow has measured it to be, and the entry sizes are derived
		/// from that (MainWindow.UpdateSlideThumbnailSize).
		///
		/// The height is likewise untouched: it follows the host, see
		/// <see cref="FollowHostHeight"/>.
		/// </summary>
		public void ApplyStyle(string listStyleKey)
		{
			Style style = Application.Current != null
				? Application.Current.TryFindResource(listStyleKey) as Style
				: null;
			if (style != null)
			{
				this.SlideListBox.Style = style;

				// The list re-templates with the style: the ScrollViewer the ring
				// follows is thrown away with the old template and the containers
				// are rebuilt. One unmoving pass re-finds the viewer, and the
				// queued retry covers the containers that arrive a beat later.
				this.UpdateSelectionRing(false);
			}
		}

		// ================= width =================

		/// <summary>
		/// Settles <see cref="ContentWidth"/> and tells the owner if it moved.
		///
		/// Only written when it really changed. This runs from a SizeChanged handler,
		/// and raising the event re-measures every entry in the list, so a "changed"
		/// that did not change anything would cost a full re-layout for nothing.
		///
		/// The first run is deferred to after the current layout pass rather than
		/// done in place: it can be reached from Loaded, when the panel is still
		/// collapsed and the list has never been measured, and reading the scrollbar
		/// at that point means reading nothing - which would size the entries as
		/// though no scrollbar were taking any width at all. Deferring puts it after
		/// the pass that Loaded triggers, and every run after the first is inline,
		/// where the list is already laid out and the answer is immediate.
		/// </summary>
		private void UpdateContentWidth()
		{
			if (this._loaded)
			{
				this.MeasureContentWidth();
				return;
			}

			if (this._measureQueued)
			{
				return;
			}

			this._measureQueued = true;
			SyntacticSugar.CurrentDispatcherBeginInvoke(
				DispatcherPriority.Loaded,
				new Action(this.MeasureContentWidth));
		}

		/// <summary>
		/// Works out <see cref="ContentWidth"/> and tells the owner if it moved.
		///
		/// The sum, from a panel width of 176 and this app's 12px scrollbar:
		///
		///   176 panel
		///   - 18 the list's Margin, 9 either side          = 158 viewport
		///   - 24 the scrollbar, taken off TWICE            = 134
		///   -  1 <see cref="EntrySlack"/>                  = 133, the entry.
		///
		/// Twice is not a slip. The scrollbar template this app ships
		/// (DefaultScrollViewerStyle in Style_\SquareCorners.xaml) lays the content
		/// out across the whole viewport and draws the scrollbar OVER the rightmost 12
		/// pixels of it, rather than in a column of its own - so the viewport does not
		/// shrink to make room for the bar, and an entry has to stop short of that
		/// strip or the bar lands on top of its right edge.
		///
		/// The second half comes from the item being centred (all three tier styles
		/// set HorizontalAlignment=Center): a centred item has equal margins, so
		/// holding the right edge clear by 12 costs 12 on the left as well. That is
		/// the price of keeping the entries centred in the panel. Anchoring them left
		/// instead would give those 12 pixels back, at the cost of every entry sitting
		/// off centre in the panel with a 9px margin on one side and 22 on the other.
		/// </summary>
		private void MeasureContentWidth()
		{
			this._loaded = true;
			this._measureQueued = false;

			// NaN rather than a small number when the panel has not been measured: an
			// entry sized from an unmeasured panel would be sized from nothing.
			double width = double.NaN;
			double scrollBar = SlideThumbnailPanel.ScrollBarAllowance;
			if (!double.IsNaN(this.ActualWidth) && this.ActualWidth > 0.0)
			{
				scrollBar = this.ScrollBarWidth();
				width = Math.Floor(this.ActualWidth)
					- 2 * SlideThumbnailPanel.ListPadding
					- 2 * scrollBar
					- EntrySlack;
			}

			// NaN counts as "changed" only when the old value was a number, and a
			// number only when the old value was NaN: two unmeasured runs in a row are
			// not a change, and firing for one would make the owner re-measure every
			// entry for nothing.
			bool changed = double.IsNaN(this.ContentWidth)
				? !double.IsNaN(width)
				: double.IsNaN(width) || Math.Abs(this.ContentWidth - width) > 0.5;
			if (!changed)
			{
				return;
			}

			this.ContentWidth = width;
			Log.WriteDebug((this.IsLeftSide ? "Left" : "Right") + " thumbnail: actual=" + this.ActualWidth + " scrollbar=" + scrollBar + " contentWidth=" + width);
			Action handler = this.ContentWidthChanged;
			if (handler != null)
			{
				handler();
			}
		}

		/// <summary>
		/// How much of the list's width the scrollbar takes, measured rather than
		/// assumed - with the allowance as the floor when it cannot be measured.
		///
		/// It has to be measured because the number is not ours: it comes from this
		/// app's own scrollbar style (SquareScrollBarStyle in Style_\SquareCorners.xaml
		/// today), and the next theme, the next display scaling or the system's
		/// scrollbar width setting can all move it. A constant written here would be
		/// right on one machine and would leave the entries covered - or needlessly
		/// short - on the next one, and a covered entry is a selection ring with a side
		/// missing.
		///
		/// Found by looking for the ScrollViewer's vertical ScrollBar rather than by
		/// the name the template gives it: the name is part of the theme's business,
		/// the orientation is not. The search is re-run every time rather than kept in
		/// a field, because the list re-templates itself whenever its tier style
		/// changes and the bar we were looking at is thrown away with the old one.
		///
		/// The bar is also hooked to its own SizeChanged, so a theme that only gives
		/// it a width once it has been laid out - or changes it later - reaches
		/// MeasureContentWidth instead of leaving the entries sized from a bar that
		/// had no width at the time.
		/// </summary>
		private double ScrollBarWidth()
		{
			ScrollBar bar = this.FindScrollBar();
			if (bar == null)
			{
				// The list's template has not been applied yet, or the bar is not in it
				// at all. Fall back rather than hand back zero: zero would size the
				// entries as though nothing were standing over the right edge of the
				// list, and the scrollbar would then land on them when it appeared.
				return SlideThumbnailPanel.ScrollBarAllowance;
			}

			return bar.ActualWidth > 0.0 ? bar.ActualWidth : SlideThumbnailPanel.ScrollBarAllowance;
		}

		/// <summary>
		/// The list's vertical scrollbar, or null when there is none to be found yet.
		/// </summary>
		private ScrollBar FindScrollBar()
		{
			ScrollBar bar = SlideThumbnailPanel.FindVerticalScrollBar(this.SlideListBox);
			if (bar == null)
			{
				return null;
			}

			if (!object.ReferenceEquals(bar, this._scrollBar))
			{
				if (this._scrollBar != null)
				{
					this._scrollBar.SizeChanged -= this.OnScrollBarSizeChanged;
				}

				this._scrollBar = bar;
				bar.SizeChanged += this.OnScrollBarSizeChanged;
			}

			return bar;
		}

		/// <summary>
		/// The scrollbar is part of the width an entry may use, so an entry has to be
		/// re-measured when it moves.
		///
		/// This cannot loop: the entry's width does not feed back into the scrollbar's
		/// own width, and MeasureContentWidth drops the event when the answer did not
		/// actually move.
		/// </summary>
		private void OnScrollBarSizeChanged(object sender, SizeChangedEventArgs e)
		{
			this.MeasureContentWidth();
		}

		/// <summary>
		/// First vertical ScrollBar in the visual tree below an element, or null.
		/// Depth first - the ScrollViewer lives inside the list's template.
		/// </summary>
		private static ScrollBar FindVerticalScrollBar(DependencyObject parent)
		{
			if (parent == null)
			{
				return null;
			}

			int count = VisualTreeHelper.GetChildrenCount(parent);
			for (int i = 0; i < count; i++)
			{
				DependencyObject child = VisualTreeHelper.GetChild(parent, i);
				ScrollBar bar = child as ScrollBar;
				if (bar != null && bar.Orientation == Orientation.Vertical)
				{
					return bar;
				}

				bar = SlideThumbnailPanel.FindVerticalScrollBar(child);
				if (bar != null)
				{
					return bar;
				}
			}

			return null;
		}

		// ================= height =================

		private void OnPanelLoaded(object sender, RoutedEventArgs e)
		{
			// The host is a Grid in row 0 of MainWindow, stretched, so its own height is
			// exactly "the space above the toolbar row". Following it is what makes the
			// panel stand on top of the toolbars and grow upward rather than needing a
			// number worked out from the screen height and the toolbar tier.
			FrameworkElement host = this.Parent as FrameworkElement;
			if (host == null)
			{
				return;
			}

			this.FollowHostHeight(host);
			host.SizeChanged += delegate
			{
				this.FollowHostHeight(host);
			};

			// The width the entries may use, for the reason documented on
			// UpdateContentWidth: it can only be measured once the list has been laid
			// out with entries in it, which is after this point.
			this.UpdateContentWidth();
		}

		private void FollowHostHeight(FrameworkElement host)
		{
			double available = host.ActualHeight - TopGap;
			if (double.IsNaN(available) || available <= 0.0)
			{
				available = MinimumCap;
			}
			else if (available < MinimumCap)
			{
				available = MinimumCap;
			}

			// Only written when it really changed: assigning MaxHeight re-measures the
			// panel, and this runs from a SizeChanged handler on the element above it.
			if (Math.Abs(this.MaxHeight - available) > 0.5)
			{
				this.MaxHeight = available;
			}
		}

		// ================= click to jump =================

		/// <summary>
		/// A press on the backing rather than on the list closes the panel.
		///
		/// Tunnelling, so this runs before the list's own handler, which is why it has
		/// to decide by geometry: the press has not been routed far enough to say
		/// anything about where it landed.
		/// </summary>
		private void OnPanelPreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
		{
			Point point = e.GetPosition(this.SlideListBox);
			bool inside = point.X >= 0.0 && point.X <= this.SlideListBox.ActualWidth
				&& point.Y >= 0.0 && point.Y <= this.SlideListBox.ActualHeight;
			if (inside)
			{
				return;
			}

			Action dismiss = this.DismissRequested;
			if (dismiss != null)
			{
				dismiss();
			}
		}

		/// <summary>
		/// Swallows the press - but not one that belongs to the scrollbar.
		///
		/// Two reasons for swallowing, and both matter on a touch panel. WPF's
		/// ListBoxItem would otherwise select on the press and take the mouse capture,
		/// which is what makes a drag on this panel jump to whatever slide the finger
		/// happened to land on. And the jump itself is done on the release, where a drag
		/// has already been claimed by TouchScroll and no longer produces one.
		///
		/// The scrollbar is the exception, and it is a real one: it lives inside the
		/// list, so a press on its thumb or its page buttons reaches this handler on the
		/// way down, and swallowing it would leave a scrollbar that cannot be dragged
		/// or clicked. Marking it handled also stops the ScrollViewer from seeing the
		/// press, so the page buttons would do nothing at all. A press that lands on
		/// one is left alone and goes to the scrollbar.
		/// </summary>
		private void OnSlideListPreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
		{
			if (SlideThumbnailPanel.IsOnScrollBar(e.OriginalSource as DependencyObject))
			{
				return;
			}

			e.Handled = true;
		}

		/// <summary>
		/// Is this element the scrollbar's, or inside it?
		///
		/// Walks up from whatever was hit. A press on the thumb reports the RepeatButton
		/// that IS the thumb, so the hit element is not the ScrollBar itself, which is
		/// why this asks about the whole chain above the press rather than about one
		/// element.
		/// </summary>
		private static bool IsOnScrollBar(DependencyObject element)
		{
			DependencyObject current = element;
			while (current != null)
			{
				if (current is ScrollBar)
				{
					return true;
				}

				current = current is Visual || current is System.Windows.Media.Media3D.Visual3D
					? VisualTreeHelper.GetParent(current)
					: LogicalTreeHelper.GetParent(current);
			}

			return false;
		}

		private void OnSlideListPreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
		{
			ListBox list = sender as ListBox;
			if (list == null)
			{
				return;
			}

			int index = SlideThumbnailPanel.IndexAt(list, e.GetPosition(list));
			if (index < 0)
			{
				return;
			}

			if (list.SelectedIndex != index)
			{
				// Changing the selection is what MainWindow listens to; letting it
				// happen here also moves the selection ring, which is the same state
				// a page change drives from the other direction.
				list.SelectedIndex = index;
				return;
			}

			// The slide that is already showing: nothing about the selection changes,
			// so nothing would fire. Ask for the move directly.
			this.RaiseSlideRequested(index + 1);
		}

		/// <summary>
		/// Which entry is under a point, or -1 if the point is not on one.
		///
		/// Hit testing, then walk up to the owning ListBoxItem. It has to be the hit
		/// test: ItemContainerGenerator can only go by index or by item, and there is no
		/// "what is at this point" on it, so the point has to be resolved by hand.
		///
		/// The point is in the list's own coordinates, so it already accounts for the
		/// list having been scrolled and for the panel carrying a translate transform
		/// while it animates.
		/// </summary>
		private static int IndexAt(ListBox list, Point point)
		{
			FrameworkElement container = FindItemContainer(list.InputHitTest(point) as DependencyObject);
			ListBoxItem item = container as ListBoxItem;
			if (item == null || item.DataContext == null)
			{
				return -1;
			}

			return list.Items.IndexOf(item.DataContext);
		}

		/// <summary>
		/// Nearest ListBoxItem at or above an element, or null. Walks the visual tree,
		/// which is the tree a hit test walks.
		/// </summary>
		private static FrameworkElement FindItemContainer(DependencyObject element)
		{
			DependencyObject current = element;
			while (current != null)
			{
				ListBoxItem item = current as ListBoxItem;
				if (item != null)
				{
					return item;
				}

				current = current is Visual || current is System.Windows.Media.Media3D.Visual3D
					? VisualTreeHelper.GetParent(current)
					: LogicalTreeHelper.GetParent(current);
			}

			return null;
		}

		private void RaiseSlideRequested(int slideNumber)
		{
			Action<int> handler = this.SlideRequested;
			if (handler != null)
			{
				handler(slideNumber);
			}
		}

		// ================= selection ring =================

		/// <summary>
		/// The ring's glide, matched to the tool plate's own
		/// (BottomArea.SlideToolHighlight): one duration and one easing for both
		/// sliding selections in this app, so they feel like the same gesture.
		/// </summary>
		private static readonly TimeSpan RingGlide = TimeSpan.FromMilliseconds(220);

		/// <summary>
		/// Name the entry template gives its root Grid (Style_\ToolbarStyle.xaml,
		/// Style.ThumbnailItem*). The ring stands on that body rather than on the
		/// ListBoxItem: the container is the body plus the gap below it, and a ring
		/// around the gap is a ring around nothing.
		/// </summary>
		private const string EntryRootName = "EntryRoot";

		/// <summary>
		/// The list's ScrollViewer, once found, so scrolling can move the ring
		/// along with the entries. Replaced whenever the list re-templates itself
		/// (a tier switch changes its Style), the same way <see cref="_scrollBar"/>
		/// is.
		/// </summary>
		private ScrollViewer _listScroller;

		/// <summary>Set while a deferred ring pass is already queued.</summary>
		private bool _ringRetryQueued;

		private void OnSlideListSelectionChanged(object sender, SelectionChangedEventArgs e)
		{
			// The glide is the point: the ring TRAVELS from the old entry to the
			// new one instead of vanishing here and appearing there.
			this.UpdateSelectionRing(true);
		}

		private void OnThumbnailPanelIsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
		{
			// Appearing or going away: the ring takes its place in one frame, so
			// the panel never opens with the ring mid-flight from wherever it
			// stood the last time it was up.
			this.UpdateSelectionRing(false);
		}

		private void OnListScrollChanged(object sender, ScrollChangedEventArgs e)
		{
			// Scrolling is already a movement and the entries are already moving
			// with it; the ring joins them in the same frame. A glide here would
			// read as the ring sliding off the entry being scrolled under it.
			this.UpdateSelectionRing(false);
		}

		/// <summary>
		/// Stands the ring on the selected entry: sized to the entry's body (its
		/// template root Grid, without the gap below) and moved to where that body
		/// stands - a glide when the selection moved, a jump for everything else
		/// (scroll, resize, the panel appearing, the first pass).
		///
		/// The ring is the panel's, not the entry's: ONE Border for the whole list,
		/// translated up and down the way the tool plate translates across the tool
		/// strip. A ring drawn inside each entry - which is what this used to do -
		/// can only fade one out and another in; it cannot travel, and a mark that
		/// travels is what makes the selection read as the same thing moving rather
		/// than two unrelated things swapping.
		///
		/// There is nowhere to stand while the selected entry is not realized: the
		/// list has it scrolled out of the window, or the container has not been
		/// generated yet (ItemsSource was just reassigned and the selection put
		/// back). The ring then hides, and one deferred pass looks again - the
		/// containers usually exist a beat after the selection does, and if they
		/// still do not, the next scroll or pick will find them.
		/// </summary>
		private void UpdateSelectionRing(bool animate)
		{
			this.HookListScroller();

			Border ring = this.SelectionRing;
			if (ring == null)
			{
				return;
			}

			object selected = this.SlideListBox.SelectedItem;
			ListBoxItem item = (selected != null)
				? this.SlideListBox.ItemContainerGenerator.ContainerFromItem(selected) as ListBoxItem
				: null;
			if (item == null)
			{
				ring.Visibility = Visibility.Collapsed;
				if (selected != null)
				{
					this.QueueRingRetry();
				}

				return;
			}

			FrameworkElement entry = item.FindName(SlideThumbnailPanel.EntryRootName) as FrameworkElement;
			Point topLeft;
			try
			{
				topLeft = item.TransformToAncestor(this.RootGrid).Transform(new Point(0.0, 0.0));
			}
			catch (InvalidOperationException)
			{
				// The container is not under RootGrid yet - mid-templating. Wait it
				// out rather than stand the ring at a point that means nothing.
				ring.Visibility = Visibility.Collapsed;
				this.QueueRingRetry();
				return;
			}

			// A ring that was hidden cannot glide: it would travel from wherever it
			// last stood, which may be an entry that is long gone. It appears where
			// it belongs instead, and glides from the next change on.
			bool wasVisible = ring.Visibility == Visibility.Visible;

			ring.Visibility = Visibility.Visible;
			ring.Width = (entry != null && entry.ActualWidth > 0.0) ? entry.ActualWidth : item.ActualWidth;
			ring.Height = (entry != null && entry.ActualHeight > 0.0) ? entry.ActualHeight : item.ActualHeight;

			TranslateTransform transform = ring.RenderTransform as TranslateTransform;
			if (transform == null)
			{
				return;
			}

			if (!animate || !wasVisible)
			{
				transform.BeginAnimation(TranslateTransform.YProperty, null);
				transform.Y = topLeft.Y;
				return;
			}

			// Pure displacement, the plate's manners again: the glide starts from
			// wherever the ring is - under an animation or not - because a
			// DoubleAnimation with no From takes the current value for its start.
			transform.BeginAnimation(TranslateTransform.YProperty,
				new DoubleAnimation(topLeft.Y, SlideThumbnailPanel.RingGlide)
				{
					EasingFunction = new CubicEase { EasingMode = EasingMode.EaseOut }
				});
		}

		/// <summary>
		/// Queues one more unmoving ring pass, for the entry whose container turns
		/// up a beat after its selection does (an ItemsSource reassignment realizes
		/// containers on a later layout pass).
		/// </summary>
		private void QueueRingRetry()
		{
			if (this._ringRetryQueued)
			{
				return;
			}

			this._ringRetryQueued = true;
			SyntacticSugar.CurrentDispatcherBeginInvoke(
				DispatcherPriority.Loaded,
				new Action(delegate
				{
					this._ringRetryQueued = false;
					this.UpdateSelectionRing(false);
				}));
		}

		/// <summary>
		/// Hooks the list's ScrollViewer to <see cref="OnListScrollChanged"/>,
		/// once, re-finding it whenever the list has re-templated. The same
		/// pattern, and the same reason, as <see cref="FindScrollBar"/>.
		/// </summary>
		private void HookListScroller()
		{
			ScrollViewer viewer = SlideThumbnailPanel.FindScrollViewer(this.SlideListBox);
			if (viewer == null || object.ReferenceEquals(viewer, this._listScroller))
			{
				return;
			}

			if (this._listScroller != null)
			{
				this._listScroller.ScrollChanged -= this.OnListScrollChanged;
			}

			this._listScroller = viewer;
			viewer.ScrollChanged += this.OnListScrollChanged;
		}

		/// <summary>
		/// First ScrollViewer in the visual tree below an element, or null. Depth
		/// first - it lives inside the list's template.
		/// </summary>
		private static ScrollViewer FindScrollViewer(DependencyObject parent)
		{
			if (parent == null)
			{
				return null;
			}

			int count = VisualTreeHelper.GetChildrenCount(parent);
			for (int i = 0; i < count; i++)
			{
				DependencyObject child = VisualTreeHelper.GetChild(parent, i);
				ScrollViewer viewer = child as ScrollViewer;
				if (viewer != null)
				{
					return viewer;
				}

				viewer = SlideThumbnailPanel.FindScrollViewer(child);
				if (viewer != null)
				{
					return viewer;
				}
			}

			return null;
		}

		// ================= metrics =================

		/// <summary>
		/// Padding around the list, matching the Margin in SlideThumbnailPanel.xaml.
		/// Named ListPadding and not Padding so it cannot be confused with - or hide -
		/// Control.Padding, which means something else entirely here.
		///
		/// Private: the outside world does not do this arithmetic any more. It asks
		/// <see cref="MinimumPanelWidth"/> instead, so there is one copy of it rather
		/// than one here and one in MainWindow that can drift apart.
		/// </summary>
		private const double ListPadding = 9;

		/// <summary>
		/// How wide a panel has to be for one entry of the given width to fit inside
		/// it without the scrollbar landing on top of the entry: the entry, the list's
		/// padding on both sides, the scrollbar twice (see MeasureContentWidth for why
		/// twice), and the pixel left against rounding.
		///
		/// Used as the floor on a panel that would otherwise match a toolbar narrower
		/// than one entry - the vertical side toolbars of Office 2013 and older. Doing
		/// the arithmetic here rather than at the call site is what keeps it from
		/// drifting away from <see cref="MeasureContentWidth"/>, which takes off the
		/// same things.
		///
		/// The scrollbar is its allowance rather than a measured width: the floor is
		/// applied before the panel exists, so there is nothing yet to measure.
		/// </summary>
		internal static double MinimumPanelWidth(double entryWidth)
		{
			return entryWidth
				+ 2 * ListPadding
				+ 2 * ScrollBarAllowance
				+ EntrySlack;
		}
	}
}