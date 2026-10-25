using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.Office.Interop.PowerPoint;
using System.Windows.Controls.Primitives;
using System.Windows.Media;
using System.Windows.Threading;
using System.Windows.Media.Animation;
using ModernWpf.Controls;
using ModernWpf.Controls.Primitives;

namespace SimplePPSX
{
	/// <summary>
	/// Center toolbar: pointer / pen / eraser / fold ... exit, all on one plate.
	///
	/// Three buttons plus a fold key and the exit cross, each with a menu behind a
	/// second press. What is left here: the two Flyouts (pen colour and eraser clear),
	/// the routing of every click into PowerPoint, and the swap between the full bar
	/// and the half bar it folds down to. Show and hide stay MainWindow's job
	/// (ToolbarAnimation).
	///
	/// The magnifier and the embedded whiteboard used to sit on this toolbar; both are
	/// gone now, along with their window classes.
	/// </summary>
	public partial class BottomArea : UserControl
	{
		/// <summary>
		/// The pen colours offered in the colour picker.
		///
		/// This is the full standard palette (the 16 basic colours plus the tints and
		/// shades PowerPoint's own ink menu shows), laid out as a grid. Kept as data
		/// rather than as sixteen hand written RadioButtons so adding a colour is a
		/// one line change and the picker cannot drift out of sync with PptOperator.
		/// </summary>
		private static readonly Color[] PenColors =
		{
			// the four that PowerPoint puts first
			Color.FromRgb(0xC0, 0x00, 0x00),	// dark red
			Color.FromRgb(0xFF, 0x00, 0x00),	// red
			Color.FromRgb(0xFF, 0xC0, 0x00),	// amber
			Color.FromRgb(0xFF, 0xFF, 0x00),	// yellow
			Color.FromRgb(0x92, 0xD0, 0x50),	// light green
			Color.FromRgb(0x00, 0xB0, 0x50),	// green
			Color.FromRgb(0x00, 0xB0, 0xF0),	// cyan
			Color.FromRgb(0x00, 0x70, 0xC0),	// blue
			Color.FromRgb(0x00, 0x20, 0x60),	// dark blue
			Color.FromRgb(0x70, 0x30, 0xA0),	// purple
			Color.FromRgb(0x00, 0xFF, 0x00),	// bright green
			Color.FromRgb(0x00, 0xFF, 0xFF),	// bright cyan
			Color.FromRgb(0xFF, 0x00, 0xFF),	// magenta
			Color.FromRgb(0xFF, 0x66, 0x00),	// orange
			// neutrals
			Color.FromRgb(0x40, 0x40, 0x40),	// dark grey
			Color.FromRgb(0x80, 0x80, 0x80),	// grey
			Color.FromRgb(0xC0, 0xC0, 0xC0),	// light grey
			Color.FromRgb(0x00, 0x00, 0x00),	// black
			Color.FromRgb(0xFF, 0xFF, 0xFF)		// white
		};

		/// <summary>
		/// Should the mouse trail stop following the pointer?
		///
		/// Only the eraser claims the input; while the pointer or the pen is active
		/// the trail keeps following. (The whiteboard used to claim it as well; it
		/// went with the whiteboard.)
		/// </summary>
		public bool IsNeedStopMouseLineTrack
		{
			get
			{
				if (this.EraserToggleButton != null && this.EraserToggleButton.IsChecked == true)
				{
					return true;
				}
				return this.SelectToggleButton != null && this.SelectToggleButton.IsChecked == false
					&& (this.PenToggleButton == null || this.PenToggleButton.IsChecked == false);
			}
		}

		private Slider _eraserSlider;

		/// <summary>
		/// The pen colour panel: an ordinary element, hidden with Visibility, built once
		/// at construction. Same arrangement as the eraser panel and for the same
		/// reasons - it used to be a Flyout popup, and a popup has to fight the
		/// full-screen click-through transparent overlay to be seen and clicked.
		/// </summary>
		public Border ColorPanel { get; private set; }

		/// <summary>
		/// The colour panel's 2px black selection frame: one tile's box that slides
		/// over the grid to the selected tile. Created once by BuildColorPanel.
		/// </summary>
		public Border SwatchIndicator { get; private set; }

		/// <summary>Is the colour panel on screen? See the eraser's flag for why.</summary>
		private bool _colorMenuOpen;

		/// <summary>
		/// The eraser menu: a thick iOS-style slider that is dragged all the way to the
		/// right to erase the ink on the current slide.
		///
		/// An ordinary element in the visual tree, built once at construction and hidden
		/// with Visibility - NOT a Flyout / Popup, which is what this used to be.
		///
		/// Why that changed
		/// ================
		/// A Flyout is a separate top-level window, and every one of its problems came
		/// from that in this app:
		///
		///   - the toolbar overlay is full-screen, click-through (WS_EX_TRANSPARENT) and
		///     transparent, and the popup had to fight all three to be seen and clicked;
		///   - ShowAt had to build the popup's window and apply its template on first
		///     use, which is where the delay on opening came from;
		///   - its Closed event could land after the next press had already reopened it,
		///     which is what left the menu un-openable;
		///   - light dismiss fired on clicks that were meant for the toolbar.
		///
		/// Keeping it in the tree means it is built, styled and measured with the rest
		/// of the toolbar, costs nothing while hidden, and its visibility is a single
		/// property this class owns outright.
		/// </summary>
		public Border EraserPanel { get; private set; }

		public BottomArea()
		{
			this.InitializeComponent();
			base.DataContext = this;

			// The switch plate is built by InitializeComponent; park it on the
			// pointer before the first frame, so the bar never has a frame with
			// no selected look on it.
			this.InitToolPlate();
			this.MoveToolSwitchPlateTo(ToolKind.Pointer, animate: false);

			this.BuildColorPanel();
			this.BuildEraserPanel();
		}

		/// <summary>
		/// Builds the eraser panel up front, so that showing it later is only a
		/// Visibility change and not a construction.
		///
		/// The slider is held in a field rather than found by name afterwards, because
		/// there is no namescope to search once the tree is built in code.
		/// </summary>
		private void BuildEraserPanel()
		{
			_eraserSlider = new Slider
			{
				Minimum = 0.0,
				Maximum = 1.0,
				// The action fires on reaching the end, so ticks would only invite a
				// click just short of it that does nothing.
				IsSnapToTickEnabled = false,
				TickFrequency = 0,
				// Stretches across the panel, which is itself as wide as the toolbar.
				// A fixed 220px left a gap on the right of a wider panel.
				HorizontalAlignment = HorizontalAlignment.Stretch,
				Style = this.TryFindResource("Style.ThickSlider") as Style,
				VerticalAlignment = VerticalAlignment.Center
			};

			this.EraserPanel = new Border
			{
				Background = new SolidColorBrush(Color.FromRgb(0xFA, 0xFA, 0xFA)),
				// #666666, the same outline the toolbars use. It was #B0B0B0, which on a
				// #FAFAFA plate is so close in value that the panel had no discernible
				// edge - which is why there was "no frame around it" and no way to tell
				// where the slider was supposed to be.
				BorderBrush = new SolidColorBrush(Color.FromRgb(0x66, 0x66, 0x66)),
				// 2px on every floating panel's edge, per the user's ruling: the
				// toolbars already ring at 2px, and a panel that grows out of the
				// toolbar with a thinner rim reads as an accident.
				BorderThickness = new Thickness(2),
				// As wide as the toolbar it grows out of, so it reads as the toolbar
				// unfolding upwards rather than as a separate box.
				HorizontalAlignment = HorizontalAlignment.Stretch,
				Padding = new Thickness(12, 10, 12, 10),
				// Hidden until the button asks for it, and it stays in the layout slot
				// so the toolbar does not reflow when it appears.
				Visibility = Visibility.Collapsed,
				Child = new StackPanel
				{
					Children =
					{
						new TextBlock
						{
							// "current slide", not "every slide": see EraseInkOnCurrentSlide.
							Text = "拖到最右清除本页笔迹",
							FontFamily = new FontFamily("Microsoft YaHei UI"),
							FontSize = 11,
							Margin = new Thickness(0, 0, 0, 8)
						},
						_eraserSlider
					}
				}
			};

			this.HookEraserSliderDrag();
		}

		/// <summary>
		/// Is this the source of a press that started on one of the three tool buttons?
		///
		/// Walks up from the event source rather than comparing against the buttons
		/// directly, because OriginalSource is the innermost element the hit test
		/// landed on - the Path of the icon, not the ToggleButton.
		/// </summary>
		internal static bool IsToolButton(DependencyObject source)
		{
			return WalksUpThrough(source, typeof(System.Windows.Controls.Primitives.ToggleButton));
		}

		/// <summary>
		/// Is this press inside either tool panel?
		///
		/// Needed for the opposite reason to IsToolButton: the eraser panel hosts the
		/// slider, and the press that starts a drag lands ON that slider and on a colour
		/// swatch. Treating either as "clicked away" would tear the panel out from under
		/// the interaction the user just started.
		/// </summary>
		internal bool IsInsideToolPanel(DependencyObject source)
		{
			return WalksUpThrough(source, typeof(Border)) &&
				((this.EraserPanel != null && WalksUpThrough(source, this.EraserPanel.GetType()))
				 || (this.ColorPanel != null && WalksUpThrough(source, this.ColorPanel.GetType())));
		}

		private static bool WalksUpThrough(DependencyObject source, Type sought)
		{
			DependencyObject current = source;
			while (current != null)
			{
				if (sought.IsInstanceOfType(current))
				{
					return true;
				}

				current = System.Windows.Media.VisualTreeHelper.GetParent(current);
			}

			return false;
		}

		// ================= pen colour picker =================

		// Ten tiles to a row: a 22px tile plus 2px margin each side is a 26px pitch,
		// and the grid is pinned to ten pitches, so nineteen colours fall into exactly
		// two rows of ten and nine.
		private const double SwatchPitch = 26.0;

		private const int ColorColumns = 10;

		private void BuildColorPanel()
		{
			// Ten to a row: the WrapPanel is pinned to ten pitches, so the grid stays
			// two rows of ten and nine whatever the toolbar's width happens to be.
			WrapPanel grid = new WrapPanel
			{
				Orientation = Orientation.Horizontal,
				Width = SwatchPitch * ColorColumns,
				HorizontalAlignment = HorizontalAlignment.Left
			};

			for (int i = 0; i < PenColors.Length; i++)
			{
				grid.Children.Add(this.BuildSwatch(PenColors[i]));
			}

			// The fluent selector: one 2px black frame over the grid that slides
			// from the old tile to the new one instead of a border appearing and
			// disappearing per tile. A Grid hosts both layers so the frame can
			// sit at exact cell coordinates measured from the top-left tile.
			// The frame is the TILE's size, not the pitch: a 26px box drawn over
			// a 22px tile reads as "too big and not on any tile", which is what
			// the user said about the older one.
			this.SwatchIndicator = new Border
			{
				Width = 22,
				Height = 22,
				BorderThickness = new Thickness(2),
				BorderBrush = Brushes.Black,
				IsHitTestVisible = false,
				Visibility = Visibility.Collapsed,
				RenderTransform = new TranslateTransform()
			};

			Grid board = new Grid
			{
				Width = SwatchPitch * ColorColumns,
				HorizontalAlignment = HorizontalAlignment.Left
			};
			board.Children.Add(grid);
			board.Children.Add(this.SwatchIndicator);

			this.ColorPanel = new Border
			{
				Background = new SolidColorBrush(Color.FromRgb(0xFA, 0xFA, 0xFA)),
				// Same outline as the eraser panel and the toolbars, for the same reason,
				// at the same 2px weight every floating panel now wears.
				BorderBrush = new SolidColorBrush(Color.FromRgb(0x66, 0x66, 0x66)),
				BorderThickness = new Thickness(2),
				Padding = new Thickness(5),
				// Fixed width, not the toolbar's: ten pitches is what keeps the grid at
				// two rows, and the toolbar itself changes width per resolution tier.
				Width = SwatchPitch * ColorColumns + 12,
				HorizontalAlignment = HorizontalAlignment.Center,
				// Hidden until the pen button asks for it.
				Visibility = Visibility.Collapsed,
				Child = board
			};

			this.MarkSelectedSwatch();
		}

		/// <summary>
		/// One colour tile: a flat square with a check mark over the selected one.
		///
		/// The check is an ordinary Path rather than a glyph, for the same reason the
		/// toolbar icons are Paths: Segoe MDL2 Assets has no tick at a size that reads,
		/// and a Path scales with the tile and stays square-cornered with everything
		/// else. The tick is white with a dark outline behind it, because half of the
		/// palette is light enough that a white tick alone would vanish into it.
		/// </summary>
		private ToggleButton BuildSwatch(Color color)
		{
			Grid tile = new Grid
			{
				Background = new SolidColorBrush(color)
			};

			System.Windows.Shapes.Path tick = new System.Windows.Shapes.Path
			{
				// A check mark: down-right stroke, then up-right.
				Data = System.Windows.Media.Geometry.Parse("M 4,9 L 7.5,12.5 L 14,5"),
				Stroke = Brushes.White,
				StrokeThickness = 2.5,
				StrokeStartLineCap = PenLineCap.Round,
				StrokeEndLineCap = PenLineCap.Round,
				Width = 18,
				Height = 18,
				Stretch = Stretch.Uniform,
				HorizontalAlignment = HorizontalAlignment.Center,
				VerticalAlignment = VerticalAlignment.Center,
				Visibility = Visibility.Collapsed
			};
			tile.Children.Add(tick);

			ToggleButton swatch = new ToggleButton
			{
				// The framework's ToggleButton template, so the tile draws as the colour
				// and nothing else - a radio button's own ring would put a second shape
				// around the square.
				Background = new SolidColorBrush(color),
				BorderThickness = new Thickness(0),
				Padding = new Thickness(0),
				Margin = new Thickness(2),
				Width = 22,
				Height = 22,
				Focusable = false,
				Template = FlatSwatchTemplate,
				Content = tile,
				ToolTip = "#" + color.R.ToString("X2") + color.G.ToString("X2") + color.B.ToString("X2"),
				Name = "SwatchTick_" + color.R.ToString("X2") + color.G.ToString("X2") + color.B.ToString("X2")
			};

			swatch.Click += this.ColorSwatch_Click;
			return swatch;
		}

		/// <summary>
		/// A template that draws the content and nothing else: no frame, no chrome, no
		/// hover plate. The tile itself is the whole control.
		/// </summary>
		private static readonly ControlTemplate FlatSwatchTemplate = BuildFlatSwatchTemplate();

		private static ControlTemplate BuildFlatSwatchTemplate()
		{
			// One presenter, no frame: the tile is the whole control.
			FrameworkElementFactory presenter = new FrameworkElementFactory(typeof(ContentPresenter));
			presenter.SetValue(ContentPresenter.ContentSourceProperty, "Content");
			presenter.SetValue(ContentPresenter.HorizontalAlignmentProperty, HorizontalAlignment.Stretch);
			presenter.SetValue(ContentPresenter.VerticalAlignmentProperty, VerticalAlignment.Stretch);

			return new ControlTemplate(typeof(ToggleButton)) { VisualTree = presenter };
		}

		/// <summary>
		/// Shows the tick on the tile whose colour matches the pen's ink.
		///
		/// Done by walking the tiles rather than by binding, because the source of truth
		/// is _penColor - the same value the resync pushes and the switch plate
		/// re-hues to - so the check always sits on the colour PowerPoint actually has.
		/// </summary>
		private void MarkSelectedSwatch()
		{
			this.MarkSelectedSwatch(false);
		}

		/// <summary>
		/// The working overload: walks the tiles, ticks the one carrying the pen's
		/// colour and parks the frame on it. Pass true after a click to glide the
		/// frame over instead of jumping it - the fluent half of item 6.
		/// </summary>
		private void MarkSelectedSwatch(bool animate)
		{
			if (this.ColorPanel == null)
			{
				return;
			}

			Color current = _penColor;
			Grid board = this.ColorPanel.Child as Grid;
			WrapPanel grid = board != null ? board.Children[0] as WrapPanel : null;
			if (grid == null)
			{
				return;
			}

			for (int i = 0; i < grid.Children.Count; i++)
			{
				ToggleButton swatch = grid.Children[i] as ToggleButton;
				if (swatch == null)
				{
					continue;
				}

				SolidColorBrush brush = swatch.Background as SolidColorBrush;
				bool selected = brush != null && brush.Color == current;

				swatch.IsChecked = selected;
				Grid tile = swatch.Content as Grid;
				if (tile != null && tile.Children.Count > 0)
				{
					System.Windows.Shapes.Path tick = tile.Children[0] as System.Windows.Shapes.Path;
					if (tick != null)
					{
						tick.Visibility = selected ? Visibility.Visible : Visibility.Collapsed;
					}
				}

				if (selected)
				{
					this.SlideSwatchIndicatorTo(i, animate);
				}
			}
		}

		/// <summary>
		/// Parks (or glides) the 2px selection frame on tile <paramref name="index"/>,
		/// measured from the ten-wide grid: column = index % 10, row = index / 10.
		/// The frame is one tile's box, so each offset carries the tile's 2px
		/// margin as well as the pitch - flush on the tile, not straddling it.
		/// Pure displacement either way - the frame slides, it never fades, the same
		/// rule the toolbar fold animations follow.
		/// </summary>
		private void SlideSwatchIndicatorTo(int index, bool animate)
		{
			if (this.SwatchIndicator == null || index < 0)
			{
				return;
			}

			TranslateTransform transform = this.SwatchIndicator.RenderTransform as TranslateTransform;
			if (transform == null)
			{
				return;
			}

			// Tile origin: one pitch per cell, plus the tile's own 2px margin.
			double x = (index % ColorColumns) * SwatchPitch + 2;
			double y = (index / ColorColumns) * SwatchPitch + 2;
			this.SwatchIndicator.Visibility = Visibility.Visible;

			if (!animate)
			{
				// Drop any running glide first, or the old animation would keep
				// overriding X/Y after the numbers are written.
				transform.BeginAnimation(TranslateTransform.XProperty, null);
				transform.BeginAnimation(TranslateTransform.YProperty, null);
				transform.X = x;
				transform.Y = y;
				return;
			}

			TimeSpan glide = TimeSpan.FromMilliseconds(200);
			CubicEase ease = new CubicEase { EasingMode = EasingMode.EaseOut };
			transform.BeginAnimation(TranslateTransform.XProperty, new DoubleAnimation(x, glide) { EasingFunction = ease });
			transform.BeginAnimation(TranslateTransform.YProperty, new DoubleAnimation(y, glide) { EasingFunction = ease });
		}

		private void ColorSwatch_Click(object sender, RoutedEventArgs e)
		{
			ToggleButton swatch = sender as ToggleButton;
			if (swatch == null)
			{
				return;
			}
			SolidColorBrush brush = swatch.Background as SolidColorBrush;
			if (brush != null)
			{
				this.ApplyPenColor(brush.Color);
				// Tick plus glide: the frame slides over to the tile just clicked.
				this.MarkSelectedSwatch(true);
			}

			// Picking a colour is the "choose an item" case: the panel goes away again,
			// and the pen stays the active tool.
			this.CloseColorMenu();
			WindowsApi.SetForegroundWindow(PptOperator.Instance.SlideShowWindowIntPtr);
		}

		/// <summary>
		/// Pushes a colour into PowerPoint and onto the plate.
		///
		/// The colour-only setter on purpose: choosing a colour must not change which
		/// tool is active. The picker is reachable from the pen button, but the
		/// resync path runs for every tool, and SetPointerColor would drag PowerPoint
		/// back to the pen from the pointer or the eraser.
		/// </summary>
		private void ApplyPenColor(Color color)
		{
			// Packing is PptOperator's job: Office stores a colour as BBGGRR, not
			// RRGGBB, and doing it here is what swapped red and blue.
			PptOperator.Instance.SetPenColor(color);
			_penColor = color;
			// Record it as already written, so the next tool switch does not send the
			// same value to PowerPoint again.
			_lastPushedPenColor = color;
			// The switch plate is the colour indicator now: with the pen active it
			// re-hues to the ink colour chosen. An animated slide takes it there -
			// the frame is already parked on the pen button, so only the colour
			// truly moves.
			this.MoveToolSwitchPlateTo(_activeTool, animate: true);
		}

		// ================= the tool switch plate =================

		/// <summary>How long one tool-to-tool slide runs. Short enough to answer
		/// immediately, long enough for the ease to read.</summary>
		private static readonly TimeSpan ToolPlateTravel = TimeSpan.FromMilliseconds(220);

		/// <summary>
		/// The brush the switching plate is filled with. Mutable on purpose: the
		/// colour animates alongside the plate's travel - ink colour to the pen,
		/// toolbar accent to the pointer or the eraser - and an animated brush
		/// cannot be frozen.
		/// </summary>
		private SolidColorBrush _toolPlateBrush;

		/// <summary>
		/// Reads the plate's colour at first use from the accent resource the old
		/// checked state painted with, so the resting look matches what the
		/// toolbar showed before the plate existed.
		/// </summary>
		private void InitToolPlate()
		{
			if (this.ToolSwitchPlate == null || _toolPlateBrush != null)
			{
				return;
			}

			SolidColorBrush resource = this.TryFindResource("SystemControlHighlightAccentBrush") as SolidColorBrush;
			_toolPlateBrush = new SolidColorBrush(resource != null ? resource.Color : Color.FromRgb(0x00, 0x78, 0xD4));
			this.ToolSwitchPlate.Background = _toolPlateBrush;

			// A tier switch resizes the strip and every square button with it;
			// the plate rides along or it ends up the wrong size and places.
			this.SelectToggleButton.SizeChanged += delegate
			{
				this.MoveToolSwitchPlateTo(_activeTool, animate: false);
			};
		}

		/// <summary>
		/// One tool button's box, measured live: the buttons bind both sides to
		/// the strip height and render as squares, so the pitch is whatever the
		/// first button measures - never a tier number restated here.
		/// </summary>
		private double ToolPlatePitch()
		{
			double pitch = this.SelectToggleButton != null ? this.SelectToggleButton.ActualWidth : 0.0;
			if (pitch <= 0.0 && this.ToolButtonStrip != null)
			{
				// Nothing measured yet; the strip's own height is what the
				// buttons are sized to.
				pitch = this.ToolButtonStrip.Height;
			}
			return pitch > 0.0 ? pitch : 54.0;
		}

		/// <summary>
		/// The colour the tool plate wears on <paramref name="tool"/>: the ink
		/// colour itself over the pen, the toolbar accent over the pointer and
		/// the eraser. This IS the colour indicator now, which is why the old
		/// 4px bar on the pen button is gone.
		/// </summary>
		private Color PlateColourForTool(ToolKind tool)
		{
			if (tool == ToolKind.Pen)
			{
				return _penColor;
			}

			SolidColorBrush accent = this.TryFindResource("SystemControlHighlightAccentBrush") as SolidColorBrush;
			return accent != null ? accent.Color : Color.FromRgb(0x00, 0x78, 0xD4);
		}

		/// <summary>A glyph colour that reads on <paramref name="plate"/>: white on
		/// anything of real depth, black over light inks like yellow or white.</summary>
		private static Color ForegroundOnPlate(Color plate)
		{
			double luminance = 0.299 * plate.R + 0.587 * plate.G + 0.114 * plate.B;
			return luminance > 160.0 ? Colors.Black : Colors.White;
		}

		/// <summary>
		/// Puts the switch plate on <paramref name="tool"/>'s button - gliding
		/// there with a non-linear ease when asked, riding through the target
		/// colour as it goes.
		///
		/// Only the three tool buttons get this: the collapse and exit keys keep
		/// their own hover looks, and they stay untouched by the plate, whose
		/// width never leaves the three squares at the start of the row.
		/// </summary>
		private void MoveToolSwitchPlateTo(ToolKind tool, bool animate)
		{
			this.InitToolPlate();
			if (this.ToolSwitchPlate == null || _toolPlateBrush == null)
			{
				return;
			}

			double pitch = this.ToolPlatePitch();
			this.ToolSwitchPlate.Width = pitch;

			TranslateTransform transform = this.ToolSwitchPlate.RenderTransform as TranslateTransform;
			if (transform == null)
			{
				return;
			}

			double x = (int)tool * pitch;
			Color colour = this.PlateColourForTool(tool);
			this.ToolSwitchPlate.Visibility = Visibility.Visible;

			if (!animate)
			{
				// Cancel any clock first, exactly as SlideSwatchIndicatorTo does:
				// a running animation keeps overriding the plain values.
				_toolPlateBrush.BeginAnimation(SolidColorBrush.ColorProperty, null);
				_toolPlateBrush.Color = colour;
				transform.BeginAnimation(TranslateTransform.XProperty, null);
				transform.X = x;
			}
			else
			{
				// Quadratic EaseInOut: the plate speeds up leaving and settles
				// into its button, the same non-linear ride the colour makes.
				System.Windows.Media.Animation.QuadraticEase ease =
					new System.Windows.Media.Animation.QuadraticEase
					{
						EasingMode = EasingMode.EaseInOut
					};
				transform.BeginAnimation(
					TranslateTransform.XProperty,
					new System.Windows.Media.Animation.DoubleAnimation(x, ToolPlateTravel)
					{
						EasingFunction = ease
					});
				_toolPlateBrush.BeginAnimation(
					System.Windows.Media.SolidColorBrush.ColorProperty,
					new System.Windows.Media.Animation.ColorAnimation(colour, ToolPlateTravel)
					{
						EasingFunction = ease
					});
			}

			// The glyph has to read on whatever plate is under it: toolbar ink
			// when the plate is elsewhere, and a white/black choice over the
			// pen's own colour.
			Brush offPlate = this.TryFindResource("ToolbarForegroundBrush") as Brush
				?? new SolidColorBrush(Color.FromRgb(0x3F, 0x3F, 0x3F));
			Brush onPlate = new SolidColorBrush(ForegroundOnPlate(colour));
			this.SelectIcon.Foreground = tool == ToolKind.Pointer ? onPlate : offPlate;
			this.PenIcon.Foreground = tool == ToolKind.Pen ? onPlate : offPlate;
			this.EraserIcon.Foreground = tool == ToolKind.Eraser ? onPlate : offPlate;
		}

		// ================= the click cycle =================

		/// <summary>
		/// Owns the whole toolbar's click behaviour, because a ToggleButton cannot
		/// express it on its own.
		///
		/// A ToggleButton flips IsChecked before it raises Click, so by the time the
		/// handler runs the button no longer knows what it looked like before the
		/// press. The three-press cycle needs exactly that, so the press is intercepted
		/// on PreviewMouseLeftButtonDown (before the toggle) and handled here:
		///
		///   press 1  the tool is not the active one -> make it active (accent), no menu
		///   press 2  the tool is active, its menu closed -> open the menu
		///   press 3  the tool is active, its menu open -> close the menu
		///
		/// Choosing an item in a menu closes it too (see the Closed handlers).
		/// The pointer has no menu, so a second press on it does nothing.
		/// </summary>
		private void ToolButton_OnPreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
		{
			e.Handled = true;   // keep the ToggleButton from toggling itself

			if (ReferenceEquals(sender, this.SelectToggleButton))
			{
				this.ActivateTool(ToolKind.Pointer);
				return;
			}
			if (ReferenceEquals(sender, this.PenToggleButton))
			{
				this.ActivateTool(ToolKind.Pen);
				return;
			}
			this.ActivateTool(ToolKind.Eraser);
		}

		private enum ToolKind
		{
			Pointer,
			Pen,
			Eraser
		}

		private ToolKind _activeTool = ToolKind.Pointer;

		/// <summary>
		/// Is the eraser menu on screen?
		///
		/// Separate from <see cref="_activeTool"/> because the eraser has three states,
		/// not two. _activeTool stays Eraser throughout - the pointer really is the
		/// eraser the whole time - and this flag says whether its menu is currently
		/// open. Encoding "menu open" as a ToolKind would mean a fourth state that
		/// PowerPoint has no equivalent for, and the polling timer in
		/// MainWindow.UpdateTimerTick reads that back every 300ms and would fight it.
		/// </summary>
		private bool _eraserMenuOpen;

		/// <summary>
		/// What one press on a tool button does, per the required behaviour.
		///
		///   press 1        switch to the tool, menu closed
		///   press 2        the same tool again: open its menu and fade the button
		///                  (the "armed" look, see ApplyEraserArmedLook)
		///   press 3        the tool is active AND its menu is open: close the menu and
		///                  go back to the plain active look, still on that tool
		///
		/// Clicking anywhere else in the toolbar while the menu is open closes the
		/// menu and leaves the tool alone (DismissToolMenus), and a completed erase
		/// switches to the pen (EraserSlider_DragCompleted).
		/// </summary>
		private void ActivateTool(ToolKind tool)
		{
			if (tool != _activeTool)
			{
				// Press 1: switch to this tool. Its menu stays closed.
				this.CloseToolMenus();
				_activeTool = tool;
				this.ApplyActiveTool();
				return;
			}

			// Already the active tool: the press belongs to its menu.
			//
			// Press 3 closes it rather than reopening, because the button is only
			// "armed" while its panel is up.
			//
			// Both panels share one slot above the toolbar, so opening either one
			// takes the other down FIRST and without animation. That ordering is the
			// fix for both panels appearing at once: the old panel was being sunk
			// over 0.3s while the new one was already rising, and two overlapping
			// translucent plates above the toolbars rendered as a mess.
			if (tool == ToolKind.Pen)
			{
				if (_colorMenuOpen)
				{
					this.CloseColorMenu();
					return;
				}

				this.OpenColorPanel();
			}
			else if (tool == ToolKind.Eraser)
			{
				if (this._eraserMenuOpen)
				{
					this.CloseEraserMenu();
					return;
				}

				this.OpenEraserPanel();
			}
		}

		/// <summary>
		/// Takes the other panel down immediately, without the sink animation.
		///
		/// Immediate because the panel being opened occupies the same place: an
		/// animated close would leave the two overlapping, both translucent, both
		/// animating. Nothing is lost by cutting the outgoing one short - it is
		/// being replaced, not dismissed.
		/// </summary>
		private void TakeDownTheOtherPanel(Border keeping)
		{
			if (this.ColorPanel != null && !object.ReferenceEquals(this.ColorPanel, keeping))
			{
				_colorMenuOpen = false;
				this.ApplyPenArmedLook(false);
				this.HidePanelNow(this.ColorPanel);
			}

			if (this.EraserPanel != null && !object.ReferenceEquals(this.EraserPanel, keeping))
			{
				_eraserMenuOpen = false;
				this.ApplyEraserArmedLook(false);
				this.HidePanelNow(this.EraserPanel);
			}
		}

		/// <summary>
		/// Pushes the active tool into PowerPoint and onto the buttons. Exactly one
		/// button is checked, so exactly one is drawn in the Accent style.
		/// </summary>
		private void ApplyActiveTool()
		{
			// Visual state first, always. Everything below is a COM call that runs on
			// this UI thread and whose duration PowerPoint decides - it refuses or
			// delays calls while a slide is animating - so the buttons must already
			// show the new tool before any of it is attempted. Doing it in this order
			// means the click is acknowledged immediately even if the COM write takes
			// the better part of a second.
			this.SelectToggleButton.IsChecked = (_activeTool == ToolKind.Pointer);
			this.PenToggleButton.IsChecked = (_activeTool == ToolKind.Pen);
			this.EraserToggleButton.IsChecked = (_activeTool == ToolKind.Eraser);

			// The switch plate is the whole selected look: it glides to the new
			// button with its colour riding along. The buttons themselves carry
			// no checked plate of their own any more - see
			// Style.ToolbarToolToggleButton - so this is the one place the
			// "which tool is it" visual is painted.
			this.MoveToolSwitchPlateTo(_activeTool, animate: true);

			switch (_activeTool)
			{
				case ToolKind.Pen:
					this.SetPptPointerType(PpSlideShowPointerType.ppSlideShowPointerPen);
					this.PushCurrentPenColor();
					break;
				case ToolKind.Eraser:
					this.SetPptPointerType(PpSlideShowPointerType.ppSlideShowPointerEraser);
					break;
				default:
					this.SetPptPointerType(PpSlideShowPointerType.ppSlideShowPointerNone);
					break;
			}
		}

		/// <summary>
		/// Writes the indicator's colour into PowerPoint.
		///
		/// Goes through the raw COM properties rather than PptOperator.SetPointerColor,
		/// which forces the pointer to Pen as a side effect of writing a colour. Going
		/// through it here meant that pressing the pointer or the eraser button, or
		/// simply letting the update timer resync, would flip PowerPoint straight back
		/// to the pen - so the mouse mode could not be left once you had been in it.
		/// </summary>
		private void PushCurrentPenColor()
		{
			Color color = _penColor;
			// Skip a write of the colour that is already in PowerPoint.
			//
			// This is the one COM call on the tool-switch path, and it runs on the UI
			// thread, so its cost is felt as lag between the click and anything the
			// user sees next - including how promptly a menu appears. PowerPoint
			// refuses or delays calls while a slide is animating, so a redundant write
			// here is a redundant chance to stall.
			//
			// The resync exists to restore a colour that PowerPoint lost, so the
			// comparison is against what was last successfully written, not against
			// PowerPoint's current value (which is what the 300ms poll reads anyway).
			if (color == _lastPushedPenColor)
			{
				return;
			}

			try
			{
				PptOperator.Instance.SetPenColor(color);
				_lastPushedPenColor = color;
			}
			catch (Exception)
			{
				// The presentation went away mid-write; the next resync will catch up.
			}
		}

		/// <summary>The colour last written into PowerPoint, or null if none yet.</summary>
		private Color? _lastPushedPenColor;

		/// <summary>
		/// The ink colour the pen carries, red like PowerPoint's own first pen.
		/// The switch plate re-hues to this whenever the pen is the active tool,
		/// which is what replaces the old colour indicator: see PlateColourForTool.
		/// </summary>
		private Color _penColor = Colors.Red;

		private void CloseToolMenus()
		{
			// Both are closed outright, without the sink animation.
			//
			// This used to route through CloseColorMenu / CloseEraserMenu, which animate.
			// Two problems came of that: the toolbars' dismiss handler can call this for
			// BOTH panels in one press, and they shared a single hide timer - so the
			// second panel overwrote the first one's slot and the first was never
			// collapsed at all. It stayed Visible at zero opacity, which is why a panel
			// flashed back into view when it was next shown.
			//
			// A dismissal is not a replacement, so the animation is not wanted anyway.
			_colorMenuOpen = false;
			this.ApplyPenArmedLook(false);
			this.HidePanelNow(this.ColorPanel);

			_eraserMenuOpen = false;
			this.ApplyEraserArmedLook(false);
			this.HidePanelNow(this.EraserPanel);
		}

		/// <summary>
		/// Takes the tool panels down with the rest of the overlay when the show ends.
		///
		/// <see cref="CloseToolMenus"/> hides them outright, which is right for a
		/// dismissal - the toolbar is what the eye is following - and wrong for an
		/// exit, where every piece on screen is leaving together: HidePanel is that
		/// same travel as an animation, one sink per panel, started on this frame with
		/// the toolbars above.
		///
		/// A pending erase goes with them. The keystroke it would send has a slideshow
		/// window to address, and by now there is no longer one: it would be logged as
		/// an error and land on nothing.
		/// </summary>
		internal void SinkToolPanelsForExit()
		{
			_colorMenuOpen = false;
			this.ApplyPenArmedLook(false);
			this.HidePanel(this.ColorPanel);

			_eraserMenuOpen = false;
			this.ApplyEraserArmedLook(false);
			this.HidePanel(this.EraserPanel);
			this.CloseEraserFlyout();

			if (this._eraseAfterSinkTimer != null)
			{
				this._eraseAfterSinkTimer.Stop();
			}
		}

		/// <summary>
		/// Hides the colour panel and puts the pen button back to its plain active look.
		/// The pen itself stays the active tool.
		/// </summary>
		private void CloseColorMenu()
		{
			_colorMenuOpen = false;
			this.ApplyPenArmedLook(false);
			this.HidePanel(this.ColorPanel);
		}

		/// <summary>
		/// Shows a panel with the toolbar's rise animation.
		///
		/// ToolbarAnimation.FloatUp already is "make it visible, then bring it up into
		/// place" over 0.5s on an ease-out curve, so both panels get it by being passed
		/// in rather than each having its own animation. The rise starts below the
		/// bottom edge of the window and does not fade, so what the eye follows is the
		/// panel coming up from behind the bar it belongs to - a panel popping into its
		/// slot is not something this path can produce any more.
		/// </summary>
		private void ShowPanel(FrameworkElement panel)
		{
			if (panel == null)
			{
				return;
			}

			this.CancelPendingHide(panel);
			this.CancelPendingArrive(panel);

			// Nothing is written before FloatUp: it makes the panel Visible, puts it at
			// full opacity and moves it below the bottom edge in the same call, so no
			// frame can catch it standing in its slot. Setting Opacity = 0 first (which
			// is what this used to do, to cover a flash the old fade-in could not) is
			// now the wrong direction: without a fade, a base of zero is a panel that
			// never appears.
			ToolbarAnimation.FloatUp(panel);

			// Insurance for the resting opacity, on a timer - see StartArriveTimer.
			this.StartArriveTimer(panel);
		}

		/// <summary>How long the sink animation runs. Matches ToolbarAnimation's exit.</summary>
		private const int PanelSinkMilliseconds = 300;




		private readonly Dictionary<FrameworkElement, DispatcherTimer> _panelHideTimers =
			new Dictionary<FrameworkElement, DispatcherTimer>();

		/// <summary>How long the rise animation runs. Matches ToolbarAnimation's entry.</summary>
		private const int PanelRiseMilliseconds = 500;

		private readonly Dictionary<FrameworkElement, DispatcherTimer> _panelArriveTimers =
			new Dictionary<FrameworkElement, DispatcherTimer>();

		private void CancelPendingArrive(FrameworkElement panel)
		{
			DispatcherTimer timer;
			if (panel != null && this._panelArriveTimers.TryGetValue(panel, out timer))
			{
				timer.Stop();
			}
		}

		/// <summary>
		/// Puts the panel at its resting opacity once the rise is over.
		///
		/// The rise no longer animates opacity - it writes the resting value before the
		/// panel can be drawn - so this is insurance rather than the thing that makes
		/// the panel visible. It earns its keep by being a timer: a timer is not
		/// something an animation can cancel, so whatever happens to the clock the
		/// panel ends the rise at full opacity. This is the entrance counterpart of the
		/// departure timer, for the same reason.
		/// </summary>
		private void StartArriveTimer(FrameworkElement panel)
		{
			DispatcherTimer timer;
			if (!this._panelArriveTimers.TryGetValue(panel, out timer))
			{
				timer = new DispatcherTimer
				{
					Interval = TimeSpan.FromMilliseconds(PanelRiseMilliseconds)
				};
				FrameworkElement watched = panel;
				timer.Tick += delegate
				{
					timer.Stop();

					// Only if the panel was not hidden again in the meantime.
					if (watched.Visibility != Visibility.Visible)
					{
						return;
					}

					watched.Opacity = 1.0;
				};
				this._panelArriveTimers.Add(panel, timer);
			}

			timer.Stop();
			timer.Start();
		}

		private void CancelPendingHide(FrameworkElement panel)
		{
			DispatcherTimer timer;
			if (panel != null && this._panelHideTimers.TryGetValue(panel, out timer))
			{
				timer.Stop();
			}
		}

		/// <summary>Collapses a panel at once, with no animation.</summary>
		private void HidePanelNow(FrameworkElement panel)
		{
			if (panel == null)
			{
				return;
			}

			this.CancelPendingHide(panel);
			this.CancelPendingArrive(panel);

			// Everything the sink animation may have left behind is reset, and Opacity
			// matters most: a panel left at zero opacity is invisible but still Visible,
			// which leaves a hole in the toolbar. The next FloatUp would write the value
			// back, but the collapse must not depend on there being a next open.
			panel.Visibility = Visibility.Collapsed;
			panel.Opacity = 1.0;

			TranslateTransform transform = panel.RenderTransform as TranslateTransform;
			if (transform != null)
			{
				transform.Y = 0;
			}
		}

		/// <summary>
		/// Sinks a panel out and then collapses it.
		///
		/// The collapse is done by a timer rather than by the animation's own Completed
		/// event, and that is the fix for "it slides down but does not hide": WPF does not
		/// raise Completed when a later animation on the same property replaces this one.
		/// A timer cannot be cancelled that way, so the hide always happens. The
		/// animation is now purely the visual part, which is why no completion callback
		/// is passed to FloatDown.
		///
		/// One timer PER PANEL. A single shared timer silently dropped one of them:
		/// whichever panel was hidden second took the slot, so the first was never
		/// collapsed and stayed on screen at zero opacity.
		/// </summary>
		private void HidePanel(FrameworkElement panel)
		{
			if (panel == null)
			{
				return;
			}

			DispatcherTimer timer;
			if (!this._panelHideTimers.TryGetValue(panel, out timer))
			{
				timer = new DispatcherTimer
				{
					Interval = TimeSpan.FromMilliseconds(PanelSinkMilliseconds)
				};
				FrameworkElement watched = panel;
				timer.Tick += delegate
				{
					timer.Stop();
					if (watched.Visibility == Visibility.Collapsed)
					{
						return;
					}

					watched.Visibility = Visibility.Collapsed;
					watched.Opacity = 1.0;

					// FloatDown leaves the render transform offset behind, and leaving it
					// there means the next FloatUp starts from a stale place.
					TranslateTransform transform = watched.RenderTransform as TranslateTransform;
					if (transform != null)
					{
						transform.Y = 0;
					}
				};
				this._panelHideTimers.Add(panel, timer);
			}

			timer.Stop();
			ToolbarAnimation.FloatDown(panel, null);
			timer.Start();

		}
		/// <summary>
		/// The "armed" look for the pen button: lighter while its panel is up, so it is
		/// visible which button the panel belongs to. Same idea as the eraser.
		/// </summary>
		private void ApplyPenArmedLook(bool armed)
		{
			this.PenToggleButton.Opacity = armed ? 0.6 : 1.0;
		}

		/// <summary>Press 2 shows the colour panel.</summary>
		private void OpenColorPanel()
		{
			if (this.ColorPanel == null)
			{
				return;
			}

			// The eraser panel goes first, and immediately.
			this.TakeDownTheOtherPanel(this.ColorPanel);

			// Re-mark before showing, so the tick is on whatever colour is live even if
			// PowerPoint changed it behind our back since the panel was built.
			this.MarkSelectedSwatch();

			_colorMenuOpen = true;
			this.ApplyPenArmedLook(true);
			this.ShowPanel(this.ColorPanel);
		}

		/// <summary>Press 2 shows the eraser panel.</summary>
		private void OpenEraserPanel()
		{
			if (this.EraserPanel == null || _eraserSlider == null)
			{
				return;
			}

			// The colour panel goes first, and immediately.
			this.TakeDownTheOtherPanel(this.EraserPanel);

			_eraserMenuOpen = true;
			this.ApplyEraserArmedLook(true);
			_eraserSlider.Value = 0;
			this.ShowPanel(this.EraserPanel);
		}

		/// <summary>
		/// Closes the eraser menu and puts the button back to its plain active look.
		/// The tool itself is untouched - this is the "press 3" and "clicked away"
		/// path, both of which return to simply being on the eraser.
		/// </summary>
		private void CloseEraserMenu()
		{
			_eraserMenuOpen = false;
			this.ApplyEraserArmedLook(false);
			this.HidePanel(this.EraserPanel);
			this.CloseEraserFlyout();
		}

		/// <summary>
		/// The second press "arms" the eraser: the button goes lighter while its menu is
		/// up, so it is visible that the menu belongs to that button rather than to the
		/// slide.
		///
		/// Implemented as Opacity on the button rather than as a second style, because
		/// IsChecked is what tells PowerPoint which tool is live and must keep meaning
		/// exactly one thing. Opacity also leaves the accent plate and the white glyph
		/// readable, whereas a lighter fill would wash the white icon out.
		/// </summary>
		private void ApplyEraserArmedLook(bool armed)
		{
			this.EraserToggleButton.Opacity = armed ? 0.6 : 1.0;
		}

		/// <summary>
		/// Clicking anywhere in the toolbar while a tool menu is open closes the menu
		/// and returns to the active tool. The tool itself does not change.
		/// </summary>
		public void DismissToolMenus()
		{
			this.CloseToolMenus();
		}


		// ================= pointer =================

		private void Select_OnClick(object sender, RoutedEventArgs e)
		{
			this.ActivateTool(ToolKind.Pointer);
		}

		// ================= pen =================

		private void Pen_OnClick(object sender, RoutedEventArgs e)
		{
			// Keyboard activation arrives here; the mouse route is intercepted earlier
			// on PreviewMouseLeftButtonDown, which is where the press cycle lives.
			this.ActivateTool(ToolKind.Pen);
		}

		/// <summary>
		/// First press switches to the pen, a second press while the pen is already
		/// active opens the colour picker.
		/// </summary>
		public void PenClick()
		{
			this.ActivateTool(ToolKind.Pen);
		}

		// ================= eraser =================

		private void Rubber_OnClick(object sender, RoutedEventArgs e)
		{
			this.ActivateTool(ToolKind.Eraser);
		}

		private void CloseEraserFlyout()
		{
			// Back to the start, and the done tick goes with it: setting Value raises
			// ValueChanged, which is what hides the tick.
			if (_eraserSlider != null)
			{
				_eraserSlider.Value = 0;
			}

			this.UpdateEraserDoneTick();
		}

		/// <summary>
		/// Detects the end of a slider drag.
		/// A Slider has no Thumb property (the Thumb hangs off the Track and has to be
		/// found through the template) and no DragCompleted either, so LostMouseCapture
		/// is used instead: the Thumb captures the mouse while dragging and releases it
		/// on mouse up, and that routed event bubbles from the Thumb up to the Slider.
		/// That way no template part name is depended upon and every drag fires once.
		/// </summary>
		private void HookEraserSliderDrag()
		{
			if (_eraserSlider == null)
			{
				return;
			}
			_eraserSlider.LostMouseCapture += delegate
			{
				this.EraserSlider_DragCompleted();
			};

			// The done tick follows the value, so it appears the moment the bar reaches
			// the end and not only when the mouse is released.
			_eraserSlider.ValueChanged += delegate
			{
				this.UpdateEraserDoneTick();
			};
		}

		/// <summary>
		/// Shows the centred black tick while the eraser slider is at its end.
		///
		/// Driven from code because "reached the end" is a comparison and a Style
		/// trigger can only match an exact value. The element is found through
		/// Template.FindName, which returns nothing until the template has been
		/// applied, so this is a no-op until the panel has been shown once.
		/// </summary>
		private void UpdateEraserDoneTick()
		{
			if (_eraserSlider == null)
			{
				return;
			}

			System.Windows.Shapes.Path tick = _eraserSlider.Template?.FindName("EraserDoneTick", _eraserSlider)
				as System.Windows.Shapes.Path;
			if (tick == null)
			{
				return;
			}

			tick.Visibility = _eraserSlider.Value >= _eraserSlider.Maximum - 0.001
				? Visibility.Visible
				: Visibility.Collapsed;
		}

		private void EraserSlider_DragCompleted()
		{
			bool reachedEnd = _eraserSlider != null
				&& _eraserSlider.Value >= _eraserSlider.Maximum - 0.001;

			if (reachedEnd)
			{
				// Dragged all the way: erase the ink on this slide.
				//
				// The label used to say "every slide" and this used to call
				// PptOperator.EraseAllInkAsync, which walked Slides[i].Shapes looking
				// for msoInk shapes and deleted them. That never worked during a show:
				// ink drawn in a slideshow does not live in the slide's Shapes
				// collection, it lives in the slideshow window's own ink layer and is
				// only committed to the deck when the show ends and is saved. Verified
				// against PowerPoint 16.0 - with the pen active and strokes drawn,
				// Slides[1].Shapes.Count stayed 0, and Presentation.InkShape threw.
				//
				// PowerPoint exposes no COM API for clearing a running show's ink, but
				// it does bind E to "erase the drawings on screen" during a
				// presentation, so that is what is sent. It clears the CURRENT slide,
				// which is why the label above now says so.
				//
				// The panel goes down FIRST - with the sink animation it has
				// everywhere else - and the ink is cleared only once that sink is
				// over: close, then erase, in that order. This used to erase here and
				// now and cut the panel out in the same frame through CloseToolMenus,
				// which hides outright with no animation, so the ink and the panel
				// vanished together with nothing to watch.
				_eraserMenuOpen = false;
				this.ApplyEraserArmedLook(false);
				this.HidePanel(this.EraserPanel);
				this.CloseEraserFlyout();

				this.StartEraseAfterSink();
				return;
			}

			// Dragged but not to the end: nothing was erased, so stay on the eraser
			// with the menu closed.
			this.CloseEraserMenu();
		}

		/// <summary>
		/// The timer carrying out a pending erase, or null when none is pending.
		///
		/// One timer, restarted rather than stacked: a second drag to the end while
		/// the first erase is still waiting just moves the erase, it does not queue
		/// another one. Stopped in InitializeBottomArea so an erase started in one
		/// presentation cannot land in the next.
		/// </summary>
		private DispatcherTimer _eraseAfterSinkTimer;

		/// <summary>
		/// Schedules the erase for the moment the eraser panel has finished sinking.
		///
		/// A timer rather than a completion callback for the reason HidePanel gives:
		/// WPF does not raise Completed when a later animation replaces this one, and
		/// an erase that never ran would leave the ink on the slide with the panel
		/// already gone and no way to try again. The interval is PanelSinkMilliseconds,
		/// so this fires on the same tick that collapses the panel - the close is over
		/// and the erase happens then, never before.
		/// </summary>
		private void StartEraseAfterSink()
		{
			if (this._eraseAfterSinkTimer == null)
			{
				this._eraseAfterSinkTimer = new DispatcherTimer
				{
					Interval = TimeSpan.FromMilliseconds(PanelSinkMilliseconds)
				};
				this._eraseAfterSinkTimer.Tick += delegate
				{
					this._eraseAfterSinkTimer.Stop();
					this.CompleteErase();
				};
			}

			this._eraseAfterSinkTimer.Stop();
			this._eraseAfterSinkTimer.Start();
		}

		/// <summary>
		/// Clears the ink on the slide that is showing, then hands the pen back.
		///
		/// The hand-back only happens while the eraser is still the live tool: if the
		/// user switched away during the sink, that choice stands. The eraser was only
		/// ever a one-shot tool - once the ink is gone there is nothing left for it to
		/// do - but it must not pull the tool back out of a hand the user just gave it
		/// to something else.
		/// </summary>
		private void CompleteErase()
		{
			Log.WriteDebug("eraser: panel down, clearing ink now");
			PptOperator.Instance.EraseInkOnCurrentSlide();

			if (_activeTool == ToolKind.Eraser)
			{
				_activeTool = ToolKind.Pen;
				this.ApplyActiveTool();
			}
		}

		/// <summary>Resets the tool state before every presentation.</summary>
		public void InitializeBottomArea()
		{
			_activeTool = ToolKind.Pointer;
			if (this.SelectToggleButton != null)
			{
				this.SelectToggleButton.IsChecked = true;
			}
			if (this.PenToggleButton != null)
			{
				this.PenToggleButton.IsChecked = false;
			}
			if (this.EraserToggleButton != null)
			{
				this.EraserToggleButton.IsChecked = false;
			}
			// Park the switch plate on the pointer without any glide: the bar is
			// rising for a new show, and a plate caught mid-travel reads as a
			// leftover from the last one.
			this.MoveToolSwitchPlateTo(ToolKind.Pointer, animate: false);
			this.CloseToolMenus();
			this.CloseEraserFlyout();

			// An erase that is still waiting out its sink must not fire into the
			// presentation that is about to start: E would go to whatever slide the
			// new show opens on.
			if (this._eraseAfterSinkTimer != null)
			{
				this._eraseAfterSinkTimer.Stop();
			}

			this.ResetToolbarBar();
		}

		// ================= collapse / expand =================

		/// <summary>
		/// How long the bar takes to sink. Matches ToolbarAnimation's exit, so a fold
		/// reads the same as the toolbar leaving the screen.
		/// </summary>
		private const int BarSinkMilliseconds = 300;

		/// <summary>
		/// The swap timer, null while the toolbar is at rest.
		///
		/// The halves are swapped on a timer rather than on the animation's Completed
		/// event, for the same reason <see cref="HidePanel"/> is: WPF does not raise
		/// Completed when a later animation replaces this one, and a swap that never
		/// runs would leave the toolbar with neither half on screen.
		/// </summary>
		private DispatcherTimer _barSwapTimer;

		/// <summary>Which half is on screen. Reset for every presentation.</summary>
		private bool _isToolbarFolded;

		/// <summary>True while the halves are swapping, so a second press waits for it.</summary>
		private bool _isBarSwapRunning;

		/// <summary>
		/// The key at the end of the toolbar: fold it down to the half bar, or bring
		/// the half bar back up to the full one. Same key on both, because both do the
		/// same thing.
		/// </summary>
		private void CollapseToggle_OnClick(object sender, RoutedEventArgs e)
		{
			this.ToggleToolbarBar();
		}

		/// <summary>
		/// Swaps between the full toolbar and the folded half bar.
		///
		/// The half bar always travels first and the full one always travels second,
		/// and each half runs the ordinary exit/entry animation - down through the
		/// bottom edge, then up out of it, with no fade either way - so the swap looks
		/// like the toolbar leaving and a shorter one arriving rather than like a jump
		/// cut.
		///
		/// The four side toolbars are part of that one movement, not a second one:
		/// folding, they sink in the same frame as the bar; unfolding, they rise in the
		/// frame the full bar rises in - which means they WAIT here for the half bar to
		/// finish sinking, because starting them now would put the two ends of the row
		/// on screen 300ms before its middle. See StartBarSwapTimer for the other half.
		///
		/// The width of the half bar is captured BEFORE the full one is hidden: a
		/// hidden element is not measured and reports zero, so reading it after the
		/// swap would fold the bar down to nothing. It is never bound for that reason.
		/// </summary>
		public void ToggleToolbarBar()
		{
			if (this._isBarSwapRunning)
			{
				return;
			}

			// The open tool menus are the panels' own thing: they are hosted in the row
			// above this bar and would be left floating over the slide once the bar they
			// came from has folded away.
			this.DismissToolMenus();

			FrameworkElement outgoing;
			FrameworkElement incoming;
			if (this._isToolbarFolded)
			{
				this._isToolbarFolded = false;
				outgoing = this.CollapsedBar;
				incoming = this.ExpandedBar;
			}
			else
			{
				// One tool button wide now, per the user's design: half as tall
				// as the bar, one button across, the glyph carrying it. The
				// buttons bind their sides to the strip height, so the key's
				// width is the strip's own height - read before the full bar is
				// hidden, for the same "hidden elements measure zero" reason the
				// old full-width capture noted.
				this.CollapsedBar.Width = this.ToolButtonStrip.ActualHeight > 0
					? this.ToolButtonStrip.ActualHeight
					: this.ToolButtonStrip.Height;
				this._isToolbarFolded = true;
				outgoing = this.ExpandedBar;
				incoming = this.CollapsedBar;
			}

			this._isBarSwapRunning = true;

			Log.WriteDebug(string.Format(
				"toolbar {0}: out={1} in={2} width={3}",
				this._isToolbarFolded ? "folding" : "unfolding",
				outgoing.Name, incoming.Name, this.CollapsedBar.Width));

			ToolbarAnimation.FloatDown(outgoing, null);

			// FOLDING: the four side toolbars go down in the same breath as the bar.
			// This is the same frame, so it reads as one movement rather than as two.
			//
			// UNFOLDING: nothing happens here on purpose. They come back with the full
			// bar instead - StartBarSwapTimer brings them up on the tick that raises it,
			// otherwise the two ends of the row would arrive 300ms ahead of its middle.
			if (this._isToolbarFolded)
			{
				MainWindow mainWindow = ViewHelper.MainWindow;
				if (mainWindow != null)
				{
					mainWindow.SetSideToolbarsFolded(true);
				}
			}

			this.StartBarSwapTimer(outgoing, incoming, !this._isToolbarFolded);
		}

		/// <param name="bringSideToolbarsBack">
		/// True when the bar being raised is the full one: the four side toolbars come
		/// back on this same tick, so the row rises as a single movement.
		/// </param>
		private void StartBarSwapTimer(FrameworkElement outgoing, FrameworkElement incoming, bool bringSideToolbarsBack)
		{
			this.StopBarSwapTimer();

			DispatcherTimer timer = new DispatcherTimer
			{
				Interval = TimeSpan.FromMilliseconds(BottomArea.BarSinkMilliseconds)
			};
			timer.Tick += delegate
			{
				this.StopBarSwapTimer();

				// FloatDown leaves the sink offset behind, and a bar parked there is
				// below the bottom edge but still Visible - so the next time it comes
				// up it would have to be dragged back before it could be seen. Reset it
				// on the way out, exactly as HidePanelNow does for the panels.
				outgoing.Visibility = Visibility.Collapsed;
				outgoing.Opacity = 1.0;
				BottomArea.ResetBarTransform(outgoing);

				// Visibility first, then the rise: FloatUp is what writes the opacity
				// and starts the bar below the bottom edge (see ToolbarAnimation), so
				// there is no frame in which the bar can be seen standing in place.
				incoming.Visibility = Visibility.Visible;
				this._isBarSwapRunning = false;
				ToolbarAnimation.FloatUp(incoming);

				// Log the state anyway: "the bar came back and answers the mouse" is the
				// one question this swap has to answer, and it is the one that has been
				// wrong before.
				Log.WriteDebug(string.Format(
					"bar swap done: {0} vis={1} opacity={2} hitTest={3}",
					incoming.Name, incoming.Visibility, incoming.Opacity,
					incoming.IsHitTestVisible));

				// The other half of the timing. The four side toolbars rise on THIS
				// tick, in the same frame as the full bar, so the whole row comes back
				// as one movement; started back at the key press they would have been
				// standing 300ms ahead of the middle of the row they belong to.
				if (bringSideToolbarsBack)
				{
					MainWindow mainWindow = ViewHelper.MainWindow;
					if (mainWindow != null)
					{
						mainWindow.SetSideToolbarsFolded(false);
					}
				}
			};

			this._barSwapTimer = timer;
			timer.Start();
		}

		private void StopBarSwapTimer()
		{
			if (this._barSwapTimer != null)
			{
				this._barSwapTimer.Stop();
				this._barSwapTimer = null;
			}
		}

		/// <summary>
		/// Puts the toolbar back to the full bar, with no animation, before a show.
		///
		/// The fold never survives a presentation: the entry animation is written for
		/// the full bar, and a timer still in flight from the last one would hide the
		/// bar the moment it appeared.
		/// </summary>
		private void ResetToolbarBar()
		{
			if (this.ExpandedBar == null || this.CollapsedBar == null)
			{
				return;
			}

			this.StopBarSwapTimer();
			this._isBarSwapRunning = false;
			this._isToolbarFolded = false;

			// Detach what is still running from the last presentation rather than let it
			// finish underneath: the reset below writes local values, which a running
			// animation ignores, and a sink left in flight would carry the full bar down
			// through the bottom edge a moment after this show started - and then write
			// that offset into the base when it completed. ResetBarTransform (below) is
			// what detaches the sink; the opacity clock goes here, so a dim left over
			// from an earlier show cannot darken the bar either.
			this.CollapsedBar.BeginAnimation(UIElement.OpacityProperty, null);
			this.ExpandedBar.BeginAnimation(UIElement.OpacityProperty, null);

			this.CollapsedBar.Visibility = Visibility.Collapsed;
			this.CollapsedBar.Opacity = 1.0;
			this.CollapsedBar.IsHitTestVisible = false;
			BottomArea.ResetBarTransform(this.CollapsedBar);

			this.ExpandedBar.Visibility = Visibility.Visible;
			this.ExpandedBar.Opacity = 1.0;
			// FloatDown takes hit testing off the half that is sinking; the full bar is
			// the one that has to answer the mouse from the first frame of the show.
			this.ExpandedBar.IsHitTestVisible = true;
			BottomArea.ResetBarTransform(this.ExpandedBar);
		}

		private static void ResetBarTransform(FrameworkElement bar)
		{
			TranslateTransform transform = bar.RenderTransform as TranslateTransform;
			if (transform != null)
			{
				// Cancel first: a running clock keeps its own value over whatever is
				// written here, and a sink still in flight would drop the bar back
				// through the bottom edge when it finished - or, through its own
				// Completed, write that off-screen offset into the base for good.
				transform.BeginAnimation(TranslateTransform.YProperty, null);
				transform.Y = 0;
			}
		}

		// ================= exit =================

		private void Exit_OnClick(object sender, RoutedEventArgs e)
		{
			this.ExitApplication();
		}

		/// <summary>
		/// The one exit route, shared by the on-screen exit button, the Escape key and
		/// the tray menu.
		/// </summary>
		public void ExitApplication()
		{
			this.EndPresentation();
		}

		/// <summary>
		/// The tray's Exit: end the presentation if one is running, then close the app.
		///
		/// Split from <see cref="ExitApplication"/> on purpose. The on-screen exit button
		/// and the Escape key only end the presentation and leave the app in the tray -
		/// that is what they have always done. The tray menu item promises to quit, so
		/// it has to close the app too, and it used to do neither reliably: PptOperator.Exit
		/// returns early when there is no slideshow window (exactly the idle state the
		/// tray sits in), so nothing closed the app at all.
		/// </summary>
		public void ExitAndShutdown()
		{
			try
			{
				this.EndPresentation();
			}
			catch (Exception)
			{
				// PowerPoint may already be gone; the app still has to close.
			}

			System.Windows.Application app = System.Windows.Application.Current;
			if (app != null)
			{
				app.Shutdown();
			}
			else
			{
				Environment.Exit(0);
			}
		}

		/// <summary>
		/// Ends the running presentation. Does not touch the lifetime of this app.
		/// </summary>
		private void EndPresentation()
		{
			// ViewHelper.MainWindow is only built by BuildViews(). With neither a
			// presentation running nor "preload window" enabled the tray's exit lands
			// here with MainWindow still null, and the old code dereferenced
			// IsExiting on it.
			MainWindow mainWindow = ViewHelper.MainWindow;
			if (mainWindow != null)
			{
				mainWindow.IsExiting = true;
			}

			PptOperator.Instance.Exit();
		}

		// ================= state sync =================

		private void SetPptPointerType(PpSlideShowPointerType type)
		{
			// Remember what was asked for, and when. See ApplyPendingToolRequest.
			_requestedPointerType = type;
			_requestedPointerTypeAt = DateTime.UtcNow;

			PptOperator.Instance.PointerType = type;
			WindowsApi.SetForegroundWindow(PptOperator.Instance.SlideShowWindowIntPtr);
		}

		/// <summary>The pointer type this toolbar last asked PowerPoint for.</summary>
		private PpSlideShowPointerType _requestedPointerType = PpSlideShowPointerType.ppSlideShowPointerNone;

		private DateTime _requestedPointerTypeAt = DateTime.MinValue;

		/// <summary>
		/// How long a tool request is given before PowerPoint's own answer wins again.
		///
		/// Short enough that a genuine change made inside PowerPoint (a keyboard
		/// shortcut, say) is still picked up promptly, long enough to cover the
		/// request never being applied at all.
		/// </summary>
		private static readonly TimeSpan ToolRequestGrace = TimeSpan.FromSeconds(2.0);

		/// <summary>
		/// Decides whether PowerPoint's reported pointer type may replace the tool the
		/// user just picked.
		///
		/// WHY THIS IS NEEDED
		/// =================
		/// This is the reason the second press on a tool button did nothing. The
		/// sequence, measured from the reported symptom ("it jumps back to the pen"):
		///
		///   press 1   ActivateTool(Eraser) sets the tool and asks PowerPoint for the
		///             eraser pointer
		///   +300ms    UpdateTimerTick calls this. PowerPoint had not applied the
		///             change yet - it is busy, or it drops the write while a slide is
		///             animating - so it still reports the pen. _activeTool became Pen
		///             again.
		///   press 2   ActivateTool(Eraser) now sees Eraser != Pen, so it treats the
		///             press as "switch to the eraser" and never reaches the branch
		///             that opens the menu.
		///
		/// So the menu was not failing to open; the tool was being taken away from
		/// under the second press before it could arrive.
		///
		/// The user's explicit choice therefore wins for a short grace period. After
		/// that, PowerPoint is believed again - including when it reports a tool the
		/// user never asked for, which is the case this method was originally written
		/// for.
		/// </summary>
		private bool ShouldAcceptReportedTool(PpSlideShowPointerType reported)
		{
			if (reported == _requestedPointerType)
			{
				// PowerPoint has caught up; nothing pending any more.
				return true;
			}

			return DateTime.UtcNow - _requestedPointerTypeAt > ToolRequestGrace;
		}

		/// <summary>
		/// Clears the tool selection. Used when the show ends.
		///
		/// This cannot go through <see cref="UpdateButtonState()"/>: that reads the
		/// pointer type from PowerPoint, which is exactly what has just gone away, so
		/// the read fails and the toolbar would keep the last tool lit while there is
		/// no slideshow behind it any more. Nothing is written to PowerPoint here -
		/// there is no window left to write to.
		/// </summary>
		public void ResetButtonState()
		{
			_activeTool = ToolKind.Pointer;
			this.SelectToggleButton.IsChecked = false;
			this.PenToggleButton.IsChecked = false;
			this.EraserToggleButton.IsChecked = false;
		}

		/// <summary>
		/// Refreshes the buttons from the pointer type PowerPoint reports, so the
		/// accent follows what PowerPoint thinks is active (it can change the pointer
		/// on its own, for instance when a shortcut is used during the show).
		/// </summary>
		public void UpdateButtonState()
		{
			// A rejected COM call must not be read as "the pointer is selected". See
			// PptOperator.TryGetPointerType: this timer runs every 300ms, so one bad
			// read while PowerPoint is animating a slide used to undo the tool the
			// user had just chosen - which is what made the pen switch feel dead and
			// stopped the second press of an eraser double click from opening its menu.
			PpSlideShowPointerType type;
			if (!PptOperator.Instance.TryGetPointerType(out type))
			{
				return;
			}

			// A menu open means the user is one press away from choosing something;
			// re-deciding the active tool underneath that press would turn the second
			// press back into a switch and the menu would never open.
			if (_colorMenuOpen || _eraserMenuOpen)
			{
				return;
			}

			// A tool the user picked a moment ago outranks what PowerPoint currently
			// reports. See ShouldAcceptReportedTool for the sequence this prevents.
			if (!this.ShouldAcceptReportedTool(type))
			{
				return;
			}

			ToolKind tool;
			if (type == PpSlideShowPointerType.ppSlideShowPointerPen)
			{
				tool = ToolKind.Pen;
			}
			else if (type == PpSlideShowPointerType.ppSlideShowPointerEraser)
			{
				tool = ToolKind.Eraser;
				if (PptAnalyzer.CurrentAnalyzeRate != AnalyzeRate.Highest)
				{
					PptAnalyzer.StopAnalyzeAllAsyn();
				}
			}
			else
			{
				tool = ToolKind.Pointer;
			}

			if (tool == _activeTool)
			{
				// Nothing changed. Writing the three IsChecked values again would be
				// harmless, but PushCurrentPenColor is a COM write and this runs three
				// times a second - there is no reason to issue it when the answer has
				// not moved.
				return;
			}

			_activeTool = tool;
			this.SelectToggleButton.IsChecked = (tool == ToolKind.Pointer);
			this.PenToggleButton.IsChecked = (tool == ToolKind.Pen);
			this.EraserToggleButton.IsChecked = (tool == ToolKind.Eraser);
			this.PushCurrentPenColor();
			// PowerPoint can move the tool on its own; the plate follows it like
			// any user-driven switch.
			this.MoveToolSwitchPlateTo(tool, animate: true);
		}
	}
}
