using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;
using System.Windows.Threading;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace SimplePPSX
{
	// Token: 0x02000038 RID: 56
	public class PptOperator
	{
		// Token: 0x17000045 RID: 69
		// (get) Token: 0x0600020C RID: 524 RVA: 0x0000A00F File Offset: 0x0000820F
		// (set) Token: 0x0600020D RID: 525 RVA: 0x0000A017 File Offset: 0x00008217
		public bool IsShowing { get; set; }

		// Token: 0x17000046 RID: 70
		// (get) Token: 0x0600020E RID: 526 RVA: 0x0000A020 File Offset: 0x00008220
		// (set) Token: 0x0600020F RID: 527 RVA: 0x0000A028 File Offset: 0x00008228
		public Color PointerColor { get; set; }

		// Token: 0x17000047 RID: 71
		// (get) Token: 0x06000210 RID: 528 RVA: 0x0000A031 File Offset: 0x00008231
		// (set) Token: 0x06000211 RID: 529 RVA: 0x0000A039 File Offset: 0x00008239
		public int SlidesCount { get; private set; }

		// Token: 0x17000048 RID: 72
		// (get) Token: 0x06000212 RID: 530 RVA: 0x0000A042 File Offset: 0x00008242
		// (set) Token: 0x06000213 RID: 531 RVA: 0x0000A04A File Offset: 0x0000824A
		public string PptFilePath { get; private set; }

		// Token: 0x17000049 RID: 73
		// (get) Token: 0x06000214 RID: 532 RVA: 0x0000A053 File Offset: 0x00008253
		// (set) Token: 0x06000215 RID: 533 RVA: 0x0000A05B File Offset: 0x0000825B
		public List<object> Slides { get; private set; }

		// Token: 0x1700004A RID: 74
		// (get) Token: 0x06000216 RID: 534 RVA: 0x0000A064 File Offset: 0x00008264
		// (set) Token: 0x06000217 RID: 535 RVA: 0x0000A06B File Offset: 0x0000826B
		public static PptOperator Instance { get; set; }

		// Token: 0x1700004B RID: 75
		// (get) Token: 0x06000218 RID: 536 RVA: 0x0000A073 File Offset: 0x00008273
		// (set) Token: 0x06000219 RID: 537 RVA: 0x0000A07B File Offset: 0x0000827B
		// Late-bound, not Presentation: see ComLate for why the hand-written
		// [ComImport] interop cannot be cast onto a WPS object. PptAnalyzer reads
		// SlideShowSettings, FullName and Slides off this, all by member name.
		public object Presentation { get; private set; }

		// Token: 0x1700004C RID: 76
		// (get) Token: 0x0600021A RID: 538 RVA: 0x0000A084 File Offset: 0x00008284
		// (set) Token: 0x0600021B RID: 539 RVA: 0x0000A08C File Offset: 0x0000828C
		public IntPtr SlideShowWindowIntPtr { get; private set; }

		// Token: 0x1700004D RID: 77
		// (get) Token: 0x0600021C RID: 540 RVA: 0x0000A095 File Offset: 0x00008295
		// (set) Token: 0x0600021D RID: 541 RVA: 0x0000A09D File Offset: 0x0000829D
		public PpSlideShowRangeType ShowRangeType { get; private set; }

		// Token: 0x1700004E RID: 78
		// (get) Token: 0x0600021E RID: 542 RVA: 0x0000A0A6 File Offset: 0x000082A6
		// (set) Token: 0x0600021F RID: 543 RVA: 0x0000A0D2 File Offset: 0x000082D2
		public PpSlideShowPointerType PointerType
		{
			get
			{
				PpSlideShowPointerType type;
				return this.TryGetPointerType(out type)
					? type
					: PpSlideShowPointerType.ppSlideShowPointerNone;
			}
			set
			{
				// WPS implements IDispatch but not the hand-written interface GUIDs, so
				// this goes through the name. The enum is passed as its int: the
				// declared type belongs to the local interop and is meaningless to a
				// server that never saw it.
				ComLate.Set(ComLate.Get(this._slideShowWindow, "View"), "PointerType", (int)value);
			}
		}

		/// <summary>
		/// Reads the pointer type, telling a failure apart from a real value.
		///
		/// WHY THIS EXISTS
		/// ==============
		/// The polling timer calls this three times a second to keep the toolbar in
		/// step with PowerPoint, and a failed read must NOT be allowed to look like
		/// "the pointer is selected".
		///
		/// PowerPoint refuses calls while it is busy (RPC_E_CALL_REJECTED shows up in
		/// debug.txt regularly - it happens while a slide is being animated in), and
		/// the old getter turned that into ppSlideShowPointerNone. None maps to
		/// ToolKind.Pointer, so a single rejected call during a page turn snapped the
		/// toolbar back to the arrow three times a second. That is what made the
		/// pen/eraser switching feel broken: the first press picked a tool, and the
		/// next 300ms tick undid it before the second press could open its menu, so
		/// the eraser menu never appeared no matter how fast the double click.
		///
		/// Returning false leaves the caller's last known state alone, which is the
		/// only sane reading of "I could not ask".
		/// </summary>
		public bool TryGetPointerType(out PpSlideShowPointerType type)
		{
			type = PpSlideShowPointerType.ppSlideShowPointerNone;

			object view = ComLate.Get(this._slideShowWindow, "View");
			if (view == null)
			{
				return false;
			}

			// Read as an int and converted here rather than letting the COM value
			// arrive as the enum: IDispatch hands back whatever the server thinks
			// the type is, and WPS has been seen to report it as a plain VARIANT,
			// which would not survive an implicit cast to the enum.
			object raw = ComLate.GetQuiet(view, "PointerType");
			if (raw == null)
			{
				return false;
			}

			int value;
			try
			{
				value = Convert.ToInt32(raw);
			}
			catch (Exception)
			{
				return false;
			}

			if (!Enum.IsDefined(typeof(PpSlideShowPointerType), value))
			{
				// A value outside the enum is not an error worth surfacing: it just
				// means "some tool we do not know about", which for this property
				// means "not a tool we act on". That is a real answer, not a
				// failure, so it is reported as such.
				return true;
			}

			type = (PpSlideShowPointerType)value;
			return true;
		}
		// ================= geometry of the showing deck =================

		/// <summary>
		/// What the aspect is assumed to be when the presentation will not say.
		/// 16:9, which is what PowerPoint has defaulted to since 2013.
		///
		/// Internal rather than private because MainWindow derives a thumbnail height
		/// from the ratio and needs the same fallback for the (practically impossible)
		/// case of the ratio coming back unusable.
		/// </summary>
		internal const double FallbackSlideAspect = 16.0 / 9.0;

		/// <summary>
		/// The showing deck's width / height, as its page setup reports it.
		///
		/// Read from the presentation rather than taken from
		/// PptAnalyzer.ThumbnailRate, which is only written while slides are being
		/// exported. The thumbnail panel needs the ratio before the first thumbnail
		/// exists, and after a cache hit there may never be one, so the export-time
		/// value cannot be relied on here.
		///
		/// Late bound and read with a zero fallback, like everything else on this
		/// object: a refused call gives 0 for both, and the check below turns that into
		/// the fallback ratio rather than into a division by zero.
		/// </summary>
		public double SlideAspectRatio
		{
			get
			{
				object pageSetup = ComLate.Get(this.Presentation, "PageSetup");
				double width = ComLate.GetInt(pageSetup, "SlideWidth");
				double height = ComLate.GetInt(pageSetup, "SlideHeight");
				if (width <= 0.0 || height <= 0.0)
				{
					return PptOperator.FallbackSlideAspect;
				}

				return width / height;
			}
		}

		// Token: 0x1700004F RID: 79
		// (get) Token: 0x06000220 RID: 544 RVA: 0x0000A0F8 File Offset: 0x000082F8
		public int CurrentShowPosition
		{
			get
			{
				object view = ComLate.Get(this._slideShowWindow, "View");
				// 1 is the fallback the typed version used, and it is a sane answer:
				// "first slide", which is where a show that is not running yet belongs.
				return ComLate.GetInt(view, "CurrentShowPosition", 1);
			}
		}

		// Token: 0x17000050 RID: 80
		// (get) Token: 0x06000221 RID: 545 RVA: 0x0000A130 File Offset: 0x00008330
		public int EndingSlide
		{
			get
			{
				return ComLate.GetInt(ComLate.Get(this.Presentation, "SlideShowSettings"), "EndingSlide");
			}
		}

		// Token: 0x17000051 RID: 81
		// (get) Token: 0x06000222 RID: 546 RVA: 0x0000A142 File Offset: 0x00008342
		public int StartingSlide
		{
			get
			{
				return ComLate.GetInt(ComLate.Get(this.Presentation, "SlideShowSettings"), "StartingSlide");
			}
		}

		// Token: 0x17000052 RID: 82
		// (get) Token: 0x06000223 RID: 547 RVA: 0x0000A154 File Offset: 0x00008354
		public PpSlideShowType? PptShowType
		{
			get
			{
				// Nullable, and kept so: the caller tells "could not read the show
				// type" apart from "speaker view", and the old code did too.
				if (this.Presentation == null)
				{
					return null;
				}
				int raw = ComLate.GetInt(ComLate.Get(this.Presentation, "SlideShowSettings"), "ShowType", -1);
				if (raw < 0)
				{
					return null;
				}
				try
				{
					return new PpSlideShowType?((PpSlideShowType)raw);
				}
				catch (ArgumentException)
				{
					return null;
				}
			}
		}

		// Token: 0x06000224 RID: 548 RVA: 0x0000A16B File Offset: 0x0000836B
		// Stops a click advancing the show, which is what this app wants: the click
		// belongs to the toolbar. msoTrue = 1, msoFalse = 0, written as plain ints
		// because the MsoTriState enum is local to the hand-written interop.
		private void SetSlideAdvanceOnClickProperty(bool advanceOnClick)
		{
			if (this.Slides == null)
			{
				return;
			}
			foreach (object slide in this.Slides)
			{
				ComLate.Set(ComLate.Get(slide, "SlideShowTransition"), "AdvanceOnClick", advanceOnClick ? 1 : 0);
			}
		}

		/// <summary>
		/// The live COM slide collection from show-begin, held until
		/// <see cref="RunDeferredShowBeginWork"/> flattens it. Set to null the
		/// moment the work is claimed.
		/// </summary>
		private object _pendingSlidesCollection;

		/// <summary>The claim lock for <see cref="RunDeferredShowBeginWork"/>.</summary>
		private readonly object _deferredWorkLock = new object();

		/// <summary>The deferred work of this show, for callers arriving later.</summary>
		private Task _deferredWorkTask;

		/// <summary>
		/// The heavy half of show-begin, moved out of AppSlideShowBegin:
		///
		///   - flattening the slide collection (one COM read per slide),
		///   - PptAnalyzer.InitPptAnalyzer (SHA256 of the whole file),
		///   - the per-slide AdvanceOnClick disable (two COM calls per slide).
		///
		/// All three used to run synchronously inside the show-begin handler,
		/// between arming the entry rise and its first frame: on a large deck
		/// that is seconds of UI-thread work, and the animation was over before
		/// a single frame of it had been drawn. The work itself is unchanged -
		/// only when it runs, off the UI thread and after the rise has had its
		/// frames to itself.
		///
		/// Called by MainWindow's rise-settle timer, and by the thumbnail panel's
		/// open path which can get there first. Claim-once and thread-safe: the
		/// first caller starts the work, and every caller's onCompleted runs
		/// once it is done. A show that deferred nothing answers immediately.
		/// </summary>
		internal void RunDeferredShowBeginWork(Action onCompleted)
		{
			if (onCompleted == null)
			{
				onCompleted = delegate { };
			}

			Task work;
			lock (this._deferredWorkLock)
			{
				if (this._pendingSlidesCollection == null && this._deferredWorkTask == null)
				{
					// Nothing deferred: either a call outside a show, or one after
					// the work has already run to completion. Either way the caller
					// is not waiting for anything.
					onCompleted();
					return;
				}
				if (this._deferredWorkTask == null)
				{
					object collection = this._pendingSlidesCollection;
					this._pendingSlidesCollection = null;
					this._deferredWorkTask = Task.Run(delegate
					{
						try
						{
							// Flattened rather than kept as the live COM collection:
							// the thumbnail analysis runs for minutes over a
							// background task while the show goes on, and a
							// collection object would go stale underneath it. The
							// background thread marshals these calls the same way
							// the analysis thread has always marshaled its own.
							this.Slides = ComLate.ToList(collection);
							PptAnalyzer.InitPptAnalyzer(this.PptFilePath);
							// A click must not advance the show: it belongs to the
							// toolbar. msoFalse = 0.
							this.SetSlideAdvanceOnClickProperty(false);
							Log.WriteDebug("DeferredShowBeginWork ok: slides=" + this.SlidesCount);
						}
						catch (Exception ex)
						{
							// The show is on screen by now; losing the hash or the
							// click-disable costs the cache and that - not the show.
							Log.WriteError("DeferredShowBeginWorkError: " + ex);
						}
					});
				}
				work = this._deferredWorkTask;
			}

			work.ContinueWith(delegate
			{
				onCompleted();
			});
		}

		// Token: 0x06000225 RID: 549 RVA: 0x0000A191 File Offset: 0x00008391
		public void PreviousPage()
		{
			if (this._stopwatch.ElapsedMilliseconds < (long)PptOperator.MinPageTurningTimerSpan)
			{
				return;
			}
			// The View is resolved BEFORE the Task is dispatched on purpose: an
			// exception thrown inside Task.Run escapes to the thread pool as an
			// unobserved exception, which the outer try/catch could not see.
			object view = ComLate.Get(this._slideShowWindow, "View");
			if (view == null)
			{
				return;
			}
			Task.Run(delegate
			{
				ComLate.Call(view, "Previous");
			});
		}

		// Token: 0x06000226 RID: 550 RVA: 0x0000A1B9 File Offset: 0x000083B9
		public void NextPage()
		{
			if (this._stopwatch.ElapsedMilliseconds < (long)PptOperator.MinPageTurningTimerSpan)
			{
				return;
			}
			object view = ComLate.Get(this._slideShowWindow, "View");
			if (view == null)
			{
				return;
			}
			Task.Run(delegate
			{
				ComLate.Call(view, "Next");
			});
		}

		/// <summary>
		/// Ends the presentation, if one is running.
		///
		/// This only ever touches PowerPoint; it never shuts this application down.
		/// The tray's Exit and the Escape key both need to close the app as well, and
		/// they do that themselves - see TrayIconService.RequestExit and
		/// BottomArea.ExitApplication. (Before, the "no slideshow window" early return
		/// below was reachable from the tray, which is exactly the state the tray is in
		/// when nothing is presenting: nothing happened and the app stayed alive.)
		/// </summary>
		// Token: 0x06000227 RID: 551 RVA: 0x0000A1E1 File Offset: 0x000083E1
		public void Exit()
		{
			// _app is nulled by TryCleanComObject() (show ended / presentation closed)
			// and by the catch in CheckPowerPointProcess(), and this method is reachable
			// from the "no PowerPoint open, user hits exit" path. The old code read
			// this._app.HWND bare, which threw NullReferenceException there and then
			// triggered App.PostErrorAndShutdown -> Restart.SimplePPSX(), i.e. an
			// endless restart loop. Hence the null guard, which ComLate also provides.
			// The application window's handle, read quietly: it is only used to raise the
			// app's own window before ending the show, so failing to get it costs
			// nothing and is not worth a log line on every exit. (Unlike
			// SlideShowWindow.HWND, Application.HWND does sit on the dispatch
			// interface, so this normally succeeds.)
			int hwnd = ComLate.GetIntQuiet(this._app, "HWND");
			if (hwnd != 0)
			{
				PptOperator.SetActive((IntPtr)hwnd);
			}

			// The slideshow window is never nulled, but its COM object can be dead by
			// the time we get here, so the View is resolved through ComLate rather than
			// dereferenced.
			object view = ComLate.Get(this._slideShowWindow, "View");
			if (view == null)
			{
				// Not an error on its own. Exit() is reachable with nothing presenting
				// (the tray's Exit and the Escape key both come through here) and the
				// window's object is usually already dead by the time a show has ended,
				// which is the state recorded in error.txt as Exit_SlideShowWindowIsNull.
				// Only worth flagging when this app still believes a show is up: then
				// something failed to tell it, and the overlay would be stranded too.
				if (this.IsShowing)
				{
					Log.WriteError("Exit: a show is still marked as running but its slideshow window will not answer");
				}
				else
				{
					Log.WriteDebug("Exit: no slideshow window, nothing to end");
				}
				return;
			}
			PptAnalyzer.StopAnalyzeAllAsyn();
			Task.Run(delegate
			{
				ComLate.Call(view, "Exit");
			});
		}

		/// <summary>
		/// Erases the ink on the slide currently being shown.
		///
		/// HOW AND WHY
		/// ===========
		/// Ink drawn during a slideshow does not belong to the slide. It lives in the
		/// slideshow window's own ink layer and is only folded into the deck's Shapes
		/// collection when the show ends and the file is saved. So the obvious
		/// implementation - walk Slides[i].Shapes for msoInk and Delete each one -
		/// finds nothing at all while a show is running. Measured against PowerPoint
		/// 16.0: pen active, strokes drawn, Slides[1].Shapes.Count stayed 0, and
		/// Presentation.InkShape threw a COMException.
		///
		/// PowerPoint publishes no COM API for clearing a running show's ink. What it
		/// does have is the documented presentation shortcut E, "erase the drawings on
		/// screen", which erases the ink on the CURRENT slide. That is what is sent
		/// here, addressed to the slideshow window specifically so the keystroke
		/// cannot land on PowerPoint's editing window instead.
		///
		/// This clears one slide, not the deck, and the eraser menu's label says so.
		/// Clearing every slide would mean paging through the show and sending E on
		/// each, which would visibly flicker the presentation.
		/// </summary>
		public void EraseInkOnCurrentSlide()
		{
			IntPtr showWindow = this.SlideShowWindowIntPtr;
			if (showWindow == IntPtr.Zero)
			{
				Log.WriteError("EraseInkOnCurrentSlide: no slideshow window");
				return;
			}

			// The keystroke must go to the slideshow, not to whichever window has
			// focus: while a toolbar floats over the slide, foreground is ours.
			WindowsApi.SetForegroundWindow(showWindow);
			System.Threading.Thread.Sleep(60);
			PptOperator.SendEraseInkKey();
			Log.WriteDebug("EraseInkOnCurrentSlide: sent E to the slideshow");
		}

		/// <summary>
		/// Types E (virtual key 0x45) with keybd_event: press, then release.
		/// </summary>
		private static void SendEraseInkKey()
		{
			const byte KeyE = 0x45;
			const uint KeyUp = 0x0002;
			WindowsApi.keybd_event(KeyE, 0, 0, UIntPtr.Zero);
			WindowsApi.keybd_event(KeyE, 0, KeyUp, UIntPtr.Zero);
		}

		/// <summary>
		/// Is this shape type ink?
		///
		/// 22 = msoInk, 23 = msoInkComment, 99 = msoInkComment on older builds.
		///
		/// Kept as a named predicate, unused by the eraser menu, so that the reason
		/// slideshow ink cannot be cleared this way stays discoverable: see
		/// EraseInkOnCurrentSlide.
		/// </summary>
		internal static bool IsInkShapeType(int shapeType)
		{
			return shapeType == 22 || shapeType == 23 || shapeType == 99;
		}

		// Token: 0x06000228 RID: 552 RVA: 0x0000A20C File Offset: 0x0000840C
		public void GotoSlide(int index)
		{
			// The next-slide event is muted across the jump on purpose: a jump fires it
			// once per crossed slide, and its handler restarts the page-turn debounce,
			// which would swallow the user's next press.
			new ComAwareEventInfo(typeof(EApplication_Event), "SlideShowNextSlide").RemoveEventHandler(this._app, new EApplication_SlideShowNextSlideEventHandler(this.OnSlideShowNextSlide));
			ComLate.CallGotoSlide(ComLate.Get(this._slideShowWindow, "View"), index);
			new ComAwareEventInfo(typeof(EApplication_Event), "SlideShowNextSlide").AddEventHandler(this._app, new EApplication_SlideShowNextSlideEventHandler(this.OnSlideShowNextSlide));
			ComLate.Call(this._slideShowWindow, "Activate");
		}

		// Token: 0x06000229 RID: 553 RVA: 0x0000A28C File Offset: 0x0000848C
		public void UpdatePointerColor()
		{
			object view = ComLate.Get(this._slideShowWindow, "View");
			int rgb = ComLate.GetInt(ComLate.Get(view, "PointerColor"), "RGB");
			// Office packs a colour as 0x00BBGGRR, so the hex is read back in that
			// order and reassembled into a WPF Color, which is R,G,B.
			string text = rgb.ToString("X2").PadLeft(6, '0');
			Color red = Colors.Red;
			try
			{
				red.R = Convert.ToByte(text.Substring(4, 2), 16);
				red.G = Convert.ToByte(text.Substring(2, 2), 16);
				red.B = Convert.ToByte(text.Substring(0, 2), 16);
			}
			catch (FormatException)
			{
				// An unreadable value leaves the previous colour in place rather than
				// reporting a colour the user never chose.
				return;
			}
			this.PointerColor = red;
		}

		// Token: 0x0600022A RID: 554 RVA: 0x0000A2A0 File Offset: 0x0000A2A0
		public void SetPointerColor(int color)
		{
			this.PointerType = PpSlideShowPointerType.ppSlideShowPointerPen;
			object view = ComLate.Get(this._slideShowWindow, "View");
			ComLate.Set(ComLate.Get(view, "PointerColor"), "RGB", color);
		}

		/// <summary>
		/// Writes the pen colour without touching the pointer type.
		///
		/// SetPointerColor also forces the pointer to Pen, which is what the pen button
		/// wants but not what a colour resync wants: reading the current colour back
		/// while the pointer or the eraser is selected would silently switch PowerPoint
		/// back to the pen, and the mouse mode could never be held. This setter is the
		/// colour-only half of it.
		/// </summary>
		public void SetPenColorOnly(int color)
		{
			object view = ComLate.Get(this._slideShowWindow, "View");
			ComLate.Set(ComLate.Get(view, "PointerColor"), "RGB", color);
		}

		/// <summary>
		/// Sets the slideshow pen colour from a WPF colour.
		///
		/// This exists because Office does NOT pack a colour as 0x00RRGGBB. It packs
		/// it the other way round, as 0x00BBGGRR - which this class's own read path
		/// documents and implements:
		///
		///     // Office packs a colour as 0x00BBGGRR, so the hex is read back in that
		///     // order and reassembled into a WPF Color, which is R,G,B.
		///     red.R = text.Substring(4, 2);   // the low byte pair
		///     red.B = text.Substring(0, 2);   // the high byte pair
		///
		/// Both callers of the int overload built the value the WPF way -
		/// (R &lt;&lt; 16) | (G &lt;&lt; 8) | B - and handed it straight to Office. So
		/// red and blue arrived swapped: picking a blue pen drew red, picking red drew
		/// blue, and green was unaffected because it sits in the middle byte. That is
		/// why the colour "was wrong" in a way that looked random.
		///
		/// Packing belongs here rather than at the call sites so there is exactly one
		/// place that knows Office's byte order. A method taking a Color cannot be got
		/// wrong by a caller; an int cannot.
		/// </summary>
		public void SetPenColor(Color color)
		{
			// Office order: blue in the high byte, red in the low one.
			this.SetPenColorOnly((color.B << 16) + (color.G << 8) + color.R);
		}

		// Token: 0x0600022B RID: 555 RVA: 0x0000A2C8 File Offset: 0x0000A2C8
		//
		// Resolves the slideshow window's HWND and its Presentation, both by name.
		// The HWND is what every SetForegroundWindow call in the UI drives off, so a
		// failure here would silently leave the toolbars unable to bring the show
		// forward. The Presentation is what the thumbnail analysis walks.
		private bool CheckMainObject()
		{
			if (this._slideShowWindow == null)
			{
				Log.WriteError("SlideShowWindowIsNull");
				return false;
			}
			this.SlideShowWindowIntPtr = PptOperator.ResolveSlideShowWindowHandle();
			this.Presentation = ComLate.Get(this._slideShowWindow, "Presentation");
			if (this.Presentation == null)
			{
				Log.WriteError("PresentationIsNull");
				return false;
			}
			return true;
		}

		/// <summary>
		/// Finds the slideshow window's HWND.
		///
		/// COM cannot supply it. SlideShowWindow.HWND exists on the vtable but is NOT on
		/// the dispatch interface, so going through IDispatch - which is what ComLate
		/// does, and what this whole file now does - fails with
		/// 0x80020003 (DISP_E_MEMBERNOTFOUND). Verified against PowerPoint 16.0.
		///
		/// That handle is not optional here. MainWindow sizes its transparent overlay to
		/// it (WindowStateHelper.UpdateWindowRect) so the toolbars sit exactly over the
		/// show, and every SetForegroundWindow call uses it to put the slideshow back in
		/// front. With IntPtr.Zero the overlay would be sized to nothing.
		///
		/// So it is found by walking the presentation process's top level windows for
		/// the slideshow's own window class. Both products use the same class:
		///   PowerPoint  screenClass   (verified, title "PowerPoint 幻灯片放映 - [deck]")
		///   WPS         screenClass   (same Office-derived shell)
		/// It is checked in preference to the editing window (PPTFrameClass), which is
		/// also visible at the same time and is NOT what we want.
		/// </summary>
		private static IntPtr ResolveSlideShowWindowHandle()
		{
			IntPtr found = WindowsApi.FindWindowByClassName("screenClass");
			if (found != IntPtr.Zero)
			{
				return found;
			}

			// No fallback to the COM property: it is known NOT to be on the dispatch
			// interface (that is the whole reason for this method), so asking would
			// just log a failure on every single show. See the notes above.
			Log.WriteDebug("ResolveSlideShowWindowHandle: no screenClass window found; "
				+ "the toolbars will not be able to sit over the slideshow");
			return IntPtr.Zero;
		}

		// Token: 0x14000009 RID: 9
		// (add) Token: 0x0600022C RID: 556 RVA: 0x0000A348 File Offset: 0x00008548
		// (remove) Token: 0x0600022D RID: 557 RVA: 0x0000A380 File Offset: 0x00008580
		public event EventHandler ShowBegin;

		// Token: 0x1400000A RID: 10
		// (add) Token: 0x0600022E RID: 558 RVA: 0x0000A3B8 File Offset: 0x000085B8
		// (remove) Token: 0x0600022F RID: 559 RVA: 0x0000A3F0 File Offset: 0x000085F0
		public event EventHandler ShowEnd;

		// Token: 0x1400000B RID: 11
		// (add) Token: 0x06000230 RID: 560 RVA: 0x0000A428 File Offset: 0x00008628
		// (remove) Token: 0x06000231 RID: 561 RVA: 0x0000A460 File Offset: 0x00008660
		public event EventHandler NextSlide;

		/// <summary>
		/// Raised whenever a presentation app has just been attached to - the process
		/// is open, it answered the liveness probe and its version is known.
		///
		/// This is "PowerPoint has been opened" as far as the overlay is concerned:
		/// with the preload switch off it is the moment the window gets built and
		/// presented, so that the first presentation only has to move the toolbars
		/// onto the screen rather than also pay for the window's first present.
		/// </summary>
		public event EventHandler PowerPointAttached;

		private void OnPowerPointAttached()
		{
			EventHandler powerPointAttached = this.PowerPointAttached;
			if (powerPointAttached == null)
			{
				return;
			}
			powerPointAttached(this, EventArgs.Empty);
		}

		// Token: 0x06000232 RID: 562 RVA: 0x0000A495 File Offset: 0x00008695
		private void OnShowBegin()
		{
			EventHandler showBegin = this.ShowBegin;
			if (showBegin == null)
			{
				return;
			}
			showBegin(null, EventArgs.Empty);
		}

		// Token: 0x06000233 RID: 563 RVA: 0x0000A4AD File Offset: 0x000086AD
		private void OnShowEnd()
		{
			EventHandler showEnd = this.ShowEnd;
			if (showEnd == null)
			{
				return;
			}
			showEnd(null, EventArgs.Empty);
		}

		// Token: 0x06000234 RID: 564 RVA: 0x0000A4C5 File Offset: 0x000086C5
		private void OnNextSlide()
		{
			EventHandler nextSlide = this.NextSlide;
			if (nextSlide == null)
			{
				return;
			}
			nextSlide(this, EventArgs.Empty);
		}

		// Token: 0x06000235 RID: 565
		[DllImport("user32.dll", EntryPoint = "SetForegroundWindow")]
		private static extern bool SetActive(IntPtr hWnd);

		// Token: 0x06000236 RID: 566 RVA: 0x0000A4E0 File Offset: 0x000086E0
		public void Initialize()
		{
			this._stopwatch = new Stopwatch();
			this._checkPowerPointTimer.Interval = TimeSpan.FromMilliseconds(2500.0);
			this._checkPowerPointTimer.Tick += this.CheckPowerPointProcess;
			this._checkPowerPointTimer.Start();
			this._clearEmptyPowerPointTimer.Interval = TimeSpan.FromMilliseconds(2500.0);
			this._clearEmptyPowerPointTimer.Tick += this.ClearEmptyPowerPointProcess;
			this.CheckPowerPointProcess();
		}

		// Token: 0x06000237 RID: 567 RVA: 0x0000A56C File Offset: 0x0000876C
		private void CheckPowerPointProcess(object sender, EventArgs e)
		{
			this.CheckPowerPointProcess();
			int times = this._times;
			this._times = times + 1;
			if (times > 10)
			{
				this._times = 0;
				GC.Collect();
				GC.WaitForPendingFinalizers();
			}
		}

		// Token: 0x06000238 RID: 568 RVA: 0x0000A5A5 File Offset: 0x000087A5
		private void ClearEmptyPowerPointProcess(object sender, EventArgs e)
		{
			this.ClearEmptyPowerPointProcess();
		}

		// Token: 0x06000239 RID: 569 RVA: 0x0000A5B0 File Offset: 0x000087B0
		private void CheckPowerPointProcess()
		{
			if (ProcessHelper.HasPowerPointProcess())
			{
				try
				{
					this.CreatePowerPointApplication();
					Log.WriteDebug("Check: attach done, _app=" + (this._app != null ? "ok" : "NULL")
						+ " initOk=" + ConfigHelper.Instance.IsInitApplicationSuccessful);
					this.InitializeControl();

					// Hand over to the cleanup timer only when there is something to
					// hand over. SetTimerEnable(true) STOPS this timer, and
					// ClearEmptyPowerPointProcess does nothing at all while _app is
					// null - so handing over after a FAILED attach retired the only
					// code that ever retries. One miss (an object the Running Object
					// Table had not published yet, or a stale entry from a server that
					// had just gone) left the app detached for the whole life of that
					// PowerPoint process. Keep looking instead: this is the same
					// state as "the process has not started yet", which the timer
					// already retried every 2.5 s.
					if (this._app != null)
					{
						this.SetTimerEnable(true);

						// PowerPoint is open and answered. The overlay preloads itself
						// on this if it has not already (see ViewHelper), so the window
						// is built and presented well before the first presentation.
						this.OnPowerPointAttached();
					}
				}
				catch (Exception ex)
				{
					// Logged, not swallowed: before this the failure was invisible, which
					// is how a toolbar that never appears can go unexplained.
					Log.WriteError("CheckPowerPointProcess: " + ex);
					this._app = null;
					ConfigHelper.Instance.IsInitApplicationSuccessful = false;
				}
			}
			this._isCheckPptStateFirstly = false;
		}

		// Token: 0x0600023A RID: 570 RVA: 0x0000A608 File Offset: 0x00008808
		private void ClearEmptyPowerPointProcess()
		{
			try
			{
				if (!ProcessHelper.HasPowerPointProcess())
				{
					this.CleanAndListen();
					return;
				}

				if (this._app == null)
				{
					return;
				}

				// "The read failed" and "the deck is empty" are two different answers,
				// and this method used to see only one of them.
				//
				// ComLate answers a failed read with a fallback instead of throwing, so
				// Presentations on a busy or already dead server came back null,
				// CountOf(null) is 0, and 0 meant "clean up". Cleaning up stopped the
				// check timer and started this one; 2.5 s later the check timer
				// attached again; and so the two timers chased each other for as long
				// as the presentation app sat there with no deck open - which is
				// exactly what debug.txt records:
				//
				//     10:33:24  PPTApplicationInitialized / attach done
				//     10:33:26  ReleaseComObject   <- the Presentations read had just failed
				//     10:33:29  PPTApplicationInitialized / attach done
				//     10:33:32  ReleaseComObject
				//     ... every ~5 s, until a deck was opened at 10:37
				//
				// Every lap detached and re-attached four event handlers through
				// ComAwareEventInfo, which is reflection into a COM connection point -
				// that is where the TargetInvocationException / TargetException /
				// InvalidCastException lines the debugger prints came from. They were
				// always caught and the app never crashed (it exits with code 0), but
				// there was no reason for them to be there at all.
				object presentations = ComLate.Get(this._app, "Presentations");
				if (presentations == null)
				{
					this._unreadableTicks++;
					if (this._unreadableTicks < PptOperator.UnreadableTicksBeforeCleanup)
					{
						// More likely busy than gone - a large deck being opened, a
						// modal dialog - and a busy server answers again on the next
						// tick. Tearing down here would start the loop above over a
						// server that is perfectly alive.
						return;
					}
					// No answer five times running (12.5 s): gone, not busy.
					Log.WriteDebug("ClearEmptyPowerPointProcess: Presentations unreadable "
						+ this._unreadableTicks + " times in a row, dropping the attachment");
					this.CleanAndListen();
					return;
				}
				this._unreadableTicks = 0;

				if (ComLate.CountOf(presentations) > 0)
				{
					return;
				}

				// The read worked and the count is genuinely 0: the app is up with
				// nothing open, which is where it parks itself on its start screen.
				// That is not a reason to detach - the attachment is what lets the next
				// SlideShowBegin arrive without going looking for the server again, and
				// releasing it only to acquire it again is what drove the loop above.
				//
				// Two exceptions keep the old behaviour. WPS leaves wpp.exe running
				// after its last presentation closes, so its attachment goes with the
				// process. And an overlay that is still up has no deck left behind it,
				// which means this app was never told the show ended - tearing down is
				// what raises that and clears the screen.
				bool isWps = PptOperator.IsWpsApplication(this._app);
				if (isWps || this.IsShowing)
				{
					// Read everything off the COM object BEFORE cleaning it.
					// CleanAndListen releases the object and nulls _app, so touching
					// this._app afterwards is a NullReferenceException.
					this.CleanAndListen();

					if (isWps)
					{
						ProcessHelper.TryKillWppProcess();
					}
				}
			}
			catch (Exception ex)
			{
				// ANY failure here means the same thing in practice: the COM object is
				// gone. It used to be a single COMException with one HRESULT, so a dying
				// server that answered with something else - RPC_E_DISCONNECTED,
				// RPC_E_CALL_REJECTED, E_POINTER - fell through to a log line and the
				// overlay was left on screen.
				Log.WriteDebug("ClearEmptyPowerPointProcess: " + ex.GetType().Name
					+ " 0x" + ex.HResult.ToString("X8") + ": " + ex.Message);
				this.CleanAndListen();
			}
		}

		/// <summary>
		/// Tells the rest of the app that the presentation is over, when it ended because
		/// the presentation application went away rather than because it said so.
		///
		/// The normal ending is the SlideShowEnd COM event, which reaches
		/// MainWindow._pptOperatorShowEnd and fades the overlay out. That event is a
		/// callback FROM the COM server, so it cannot arrive if the server has already
		/// died: PowerPoint being killed, crashing, or closed with "last presentation"
		/// all leave the overlay on screen with nothing left to talk to. That is the
		/// "PPT exited and the program never closed its window" report.
		///
		/// So the cleanup path raises the same event itself. It is raised once, guarded
		/// by whether a show was actually in progress, so it cannot double-fire on top
		/// of a genuine SlideShowEnd.
		/// </summary>
		private void NotifyShowEndedUnexpectedly()
		{
			if (!this.IsShowing)
			{
				return;
			}

			this.IsShowing = false;
			Log.WriteDebug("presentation application went away: hiding the overlay");
			this.OnShowEnd();
		}

		// Token: 0x0600023B RID: 571 RVA: 0x0000A6E4 File Offset: 0x000088E4
		private void SetTimerEnable(bool isPowerPointProcessStart)
		{
			if (isPowerPointProcessStart)
			{
				this._checkPowerPointTimer.Stop();
				this._clearEmptyPowerPointTimer.Start();
				return;
			}
			this._clearEmptyPowerPointTimer.Stop();
			this._checkPowerPointTimer.Start();
		}

		private void InitializeControl()
		{
			if (this._app == null || this.IsShowing)
			{
				return;
			}
			// Picks up a show that was already running when this app started, which is
			// the normal case: the user opens a deck, presses F5, and only then this
			// toolbar appears. SlideShowWindows(1) is the presentation slide show,
			// (2) would be the rehearsal recording.
			object windows = ComLate.Get(this._app, "SlideShowWindows");
			if (ComLate.CountOf(windows) > 0)
			{
				object window = ComLate.Item(windows, 1);
				if (window == null)
				{
					return;
				}
				if (this._isCheckPptStateFirstly)
				{
					// The very first check is made while the user may still be in the
					// PowerPoint editing window, not a show. Speaker view there means
					// "editor, not presenting", and attaching would be wrong.
					object settings = ComLate.Get(ComLate.Get(this._app, "ActivePresentation"), "SlideShowSettings");
					int showType = ComLate.GetInt(settings, "ShowType", -1);
					if (showType == (int)PpSlideShowType.ppShowTypeSpeaker)
					{
						return;
					}
					this.AppSlideShowBegin(window);
					return;
				}
				this.AppSlideShowBegin(window);
			}
		}

		// Token: 0x0600023D RID: 573 RVA: 0x0000A7AC File Offset: 0x000089AC
		private void CreatePowerPointApplication()
		{
			if (this._app != null)
			{
				new ComAwareEventInfo(typeof(EApplication_Event), "SlideShowBegin").RemoveEventHandler(this._app, new EApplication_SlideShowBeginEventHandler(this.AppSlideShowBegin));
				new ComAwareEventInfo(typeof(EApplication_Event), "SlideShowEnd").RemoveEventHandler(this._app, new EApplication_SlideShowEndEventHandler(this.OnAppSlideShowEnd));
				new ComAwareEventInfo(typeof(EApplication_Event), "SlideShowNextSlide").RemoveEventHandler(this._app, new EApplication_SlideShowNextSlideEventHandler(this.OnSlideShowNextSlide));
			}
			// Assigned as a plain object, with no cast. This used to be
			// (Application)AttachToRunningApp(), and that cast was the whole reason
			// WPS support could not be made to work: see ComLate. A null here is fine -
			// every reader tolerates it, and CheckPowerPointProcess retries on the
			// timer, so a failed attach is not fatal, only delayed.
			this._app = PptOperator.AttachToRunningApp();

			if (this._app == null)
			{
				ConfigHelper.Instance.IsInitApplicationSuccessful = false;
				// Logged every tick, which is noisy but deliberate: the timer runs every
				// 2.5s and the only question this answers is "did we ever find the
				// presentation app at all", which is otherwise unanswerable.
				Log.WriteDebug("CreatePowerPointApplication: nothing registered under "
					+ string.Join(", ", PptOperator.PresentationProgIds));
				return;
			}

			// Ask it something before wiring anything to it.
			//
			// The Running Object Table keeps its entry after the server behind it has
			// gone, and such an entry hands back an object that looks perfectly good:
			// nothing about the attach itself says whether anyone is home. Wiring
			// events onto one is where the InvalidCastException in error.txt came
			// from - ComAwareEventInfo reflects into the connection point, the
			// QueryInterface for IConnectionPointContainer goes out over RPC, and the
			// answer comes back "the RPC server is unavailable" (0x800706BA).
			//
			// So probe first: Version, Path and Presentations are the three reads
			// everything else in this class already depends on, and one answer is
			// proof of life. No answer means a husk - drop it and let the check timer
			// try again rather than half-wiring events onto a corpse.
			if (!PptOperator.AppAnswers(this._app))
			{
				Log.WriteDebug("CreatePowerPointApplication: the object answered nothing and is being dropped");
				this._app = null;
				ConfigHelper.Instance.IsInitApplicationSuccessful = false;
				return;
			}

			// The version flags are only worth filling in once there is an object to
			// ask. Without this call IsWps and IsOffice2016OrLater both stayed false
			// for the life of the process: the former because nothing ever detected WPS,
			// the latter because the horizontal toolbars were never selected.
			SyntacticSugar.TryCatchAction(delegate
			{
				PptVersionManager.UpdatePptApplicationInfo(ComLate.GetString(this._app, "Path"), ComLate.GetString(this._app, "Version"));
			}, null);

			new ComAwareEventInfo(typeof(EApplication_Event), "SlideShowBegin").AddEventHandler(this._app, new EApplication_SlideShowBeginEventHandler(this.AppSlideShowBegin));
			new ComAwareEventInfo(typeof(EApplication_Event), "SlideShowEnd").AddEventHandler(this._app, new EApplication_SlideShowEndEventHandler(this.OnAppSlideShowEnd));
			new ComAwareEventInfo(typeof(EApplication_Event), "SlideShowNextSlide").AddEventHandler(this._app, new EApplication_SlideShowNextSlideEventHandler(this.OnSlideShowNextSlide));
			new ComAwareEventInfo(typeof(EApplication_Event), "PresentationCloseFinal").AddEventHandler(this._app, new EApplication_PresentationCloseFinalEventHandler(this.AppOnPresentationCloseFinal));
			ConfigHelper.Instance.IsInitApplicationSuccessful = true;
			Log.WriteDebug("PPTApplicationInitialized");
		}

		/// <summary>
		/// Does this COM object answer at all - one of Version, Path or Presentations?
		///
		/// The liveness probe behind CreatePowerPointApplication's decision to wire
		/// events or drop the object. All three are read through ComLate's quiet
		/// form, so a miss costs nothing but a first-chance exception and no log
		/// line; a null app answers false as well, since GetQuiet returns null for
		/// a null target.
		///
		/// Deliberately not a single "does Presentations read" test: a member missing
		/// from one build or another is normal, and the object is only a husk when it
		/// cannot answer anything.
		/// </summary>
		private static bool AppAnswers(object app)
		{
			return ComLate.GetQuiet(app, "Version") != null
				|| ComLate.GetQuiet(app, "Path") != null
				|| ComLate.GetQuiet(app, "Presentations") != null;
		}

		// ProgIDs to look for in the Running Object Table, in order.
		//
		// "PowerPoint.Application" is the important one and comes first: WPS
		// Presentation registers ITSELF under that ProgID when it runs, which is why
		// it is worth asking for the Microsoft ProgID to get a WPS object back. That
		// is documented behaviour rather than a guess - an AutoIt adapter for WPS
		// notes that attaching to a running WPS Presentation can return a "Microsoft
		// Office application object", i.e. the two share the entry.
		//
		// The KWPP spellings are only needed to CREATE an instance. They are kept as
		// a fallback for the case where a WPS version does not register the shared
		// entry, because without one of them the app would simply never attach on such
		// a machine. Case variants are included because ProgID lookup is
		// case-insensitive in the registry but the ROT entry's case is whatever the
		// server registered, and a CO-WPSppt driver in the wild tries the same list.
		private static readonly string[] PresentationProgIds = new string[]
		{
			"PowerPoint.Application",
			"KWPP.Application",
			"Kwpp.Application",
			"wpp.Application"
		};

		/// <summary>
		/// Attaches to the presentation application that is already running, WPS or
		/// PowerPoint, or returns null if neither is registered and running.
		///
		/// The returned object is deliberately NOT cast to anything - see ComLate for
		/// why a cast cannot work on WPS. A null return is not an error: the caller
		/// keeps _app null, every reader tolerates that, and CheckPowerPointProcess
		/// runs again on a timer, so the attach is simply retried.
		/// </summary>
		private static object AttachToRunningApp()
		{
			// WPS's own ProgID first when wpp.exe is the running process, because a
			// machine can have both installed and only the one actually running
			// matters. On a WPS-only machine that also sidesteps any PowerPoint
			// registration that is present but not serving anything.
			if (ProcessHelper.HasWpsProcess())
			{
				object wps = ComObject.TryGetActiveObject("KWPP.Application");
				if (wps != null)
				{
					return wps;
				}
			}

			for (int i = 0; i < PptOperator.PresentationProgIds.Length; i++)
			{
				object app = ComObject.TryGetActiveObject(PptOperator.PresentationProgIds[i]);
				if (app != null)
				{
					return app;
				}
			}

			return null;
		}

		/// <summary>
		/// Is this COM object a WPS Presentation rather than PowerPoint?
		///
		/// Checked off the live object rather than off PptVersionManager, because this
		/// runs on the cleanup path where the version flags may not have been filled
		/// in yet. Both the caption and the path are read defensively: they come from
		/// an object that may already be shutting down.
		/// </summary>
		private static bool IsWpsApplication(object app)
		{
			if (app == null)
			{
				return false;
			}
			try
			{
				// Late bound by name for the same reason as everything else: the
				// Caption and Path reads used to be typed property accesses, which is
				// exactly the access that cannot be cast onto a WPS object.
				string caption = ComLate.GetString(app, "Caption");
				if (!string.IsNullOrEmpty(caption) && caption.ToLowerInvariant().Contains("wps"))
				{
					return true;
				}
				string path = ComLate.GetString(app, "Path");
				if (!string.IsNullOrEmpty(path))
				{
					string lower = path.ToLowerInvariant();
					if (lower.Contains("wps") || lower.Contains("wpp") || lower.Contains("kingsoft"))
					{
						return true;
					}
				}
			}
			catch (COMException)
			{
				// The application went away while being inspected: treat as not WPS.
				return false;
			}
			catch (Exception ex)
			{
				Log.WriteError(ex.Message);
				return false;
			}
			return false;
		}

		// Token: 0x0600023E RID: 574 RVA: 0x0000A938 File Offset: 0x00008B38
		private void AppOnPresentationCloseFinal(object pres)
		{
			// Count 0 means every presentation closed, 1 means this was the last one.
			// Either way there is nothing left to drive a toolbar for.
			int open = ComLate.CountOf(ComLate.Get(this._app, "Presentations"));
			if (open <= 1)
			{
				// Only for the presentation that is actually going away: the event also
				// fires for others, and detaching then would cut the app off from a
				// show that is still running.
				if (ComLate.SameObject(ComLate.Get(this._app, "ActivePresentation"), pres))
				{
					global::System.Windows.Application application = global::System.Windows.Application.Current;
					if (application != null)
					{
						application.Dispatcher.InvokeAsync(new Action(this.CleanAndListen));
					}
				}
			}
		}

		// Token: 0x0600023F RID: 575 RVA: 0x0000A9BC File Offset: 0x00008BBC
		private void CleanAndListen()
		{
			// Tell the views first, while the state that describes the show is still
			// intact: OnShowEnd only has to raise the event, and MainWindow does the
			// hiding. Doing it before the release also means the overlay is on its way
			// out even if the release below misbehaves.
			this.NotifyShowEndedUnexpectedly();

			this.TryCleanComObject();
			this.SetTimerEnable(false);
		}

		// Token: 0x06000240 RID: 576 RVA: 0x0000A9CC File Offset: 0x00008BCC
		private void TryCleanComObject()
		{
			if (this._app == null)
			{
				return;
			}

			object app = this._app;
			// Cleared first and unconditionally. Everything below can throw, and if
			// the field were only cleared at the end then any failure would leave a
			// dead COM pointer in _app for the next caller to try again.
			this._app = null;

			GC.Collect();
			GC.WaitForPendingFinalizers();

			// Each step is separate, and each one reports on its own.
			//
			// This used to be one try block, so the FIRST failure skipped everything
			// after it: a dead event handler meant the two ReleaseComObject calls never
			// ran and the reference stayed alive. It also called ReleaseComObject
			// TWICE on the same object, which over-releases it - the second call
			// throws ArgumentException on an already-released pointer.
			//
			// RemoveEventHandler is the step that actually throws in practice: it goes
			// through ComAwareEventInfo, which is reflection-based, so a COM object
			// that has already gone away raises TargetInvocationException. That is
			// exactly the exception reported on exit, and it arrived TWICE because
			// CleanAndListen runs once when the show ends and once when the
			// application is torn down.
			DetachEvent(app, "SlideShowBegin", new EApplication_SlideShowBeginEventHandler(this.AppSlideShowBegin));
			DetachEvent(app, "SlideShowEnd", new EApplication_SlideShowEndEventHandler(this.OnAppSlideShowEnd));
			DetachEvent(app, "SlideShowNextSlide", new EApplication_SlideShowNextSlideEventHandler(this.OnSlideShowNextSlide));

			ReleaseOnce(app);
		}

		/// <summary>
		/// Removes one COM event handler, reporting a failure instead of propagating it.
		/// The exception is unwrapped: ComAwareEventInfo invokes through reflection, so
		/// what arrives is a TargetInvocationException wrapping the real COM error, and
		/// only the inner message says anything useful.
		/// </summary>
		private static void DetachEvent(object app, string eventName, Delegate handler)
		{
			try
			{
				new ComAwareEventInfo(typeof(EApplication_Event), eventName).RemoveEventHandler(app, handler);
			}
			catch (Exception ex)
			{
				Exception real = ex;
				while (real is TargetInvocationException && real.InnerException != null)
				{
					real = real.InnerException;
				}
				Log.WriteDebug(string.Format(
					"detach {0}: {1}: {2}", eventName, real.GetType().Name, real.Message));
			}
		}

		/// <summary>
		/// Releases a COM reference exactly once. A second ReleaseComObject on the
		/// same pointer throws ArgumentException, which is why this guards on the
		/// reference count rather than trusting the caller to have released already.
		/// </summary>
		private static void ReleaseOnce(object comObject)
		{
			if (comObject == null || !Marshal.IsComObject(comObject))
			{
				return;
			}

			try
			{
				// ReleaseComObject expects the reference count to be positive; reading
				// it first means an over-release is skipped rather than thrown.
				if (Marshal.ReleaseComObject(comObject) == 0)
				{
					Log.WriteDebug("ReleaseComObject: reference count was already 0");
				}
			}
			catch (Exception ex)
			{
				Log.WriteDebug("ReleaseComObject: " + ex.GetType().Name + ": " + ex.Message);
			}
		}

		// The parameter is the slideshow window the server hands us. It is late bound
		// (object) for the same reason _app is: a typed SlideShowWindow parameter would
		// make the runtime marshal the incoming pointer as the local interface, which
		// fails on WPS. The handler bodies are otherwise unchanged.
		private void AppSlideShowBegin(object wn)
		{
			try
			{
				if (ConfigHelper.Instance.IsNeedStart())
				{
					this._slideShowWindow = wn;
					if (this.CheckMainObject())
					{
						PpSlideShowType? pptShowType = this.PptShowType;
						this._isShowTypeSpeaker = pptShowType.GetValueOrDefault() == PpSlideShowType.ppShowTypeSpeaker && pptShowType != null;
						ConfigHelper.Instance.IsPptShowTypeSpeaker = this._isShowTypeSpeaker;
						if (this._isShowTypeSpeaker)
						{
							object settings = ComLate.Get(this.Presentation, "SlideShowSettings");
							this.ShowRangeType = ReadRangeType(settings);
							object slides = ComLate.Get(this.Presentation, "Slides");
							this.SlidesCount = ComLate.CountOf(slides);
							this.PptFilePath = ComLate.GetString(this.Presentation, "FullName");
							// The heavy half of show-begin is deferred to
							// RunDeferredShowBeginWork instead of done here:
							// flattening the collection (one COM read per slide),
							// the file hash in InitPptAnalyzer and the per-slide
							// click-disable together are seconds of UI-thread work
							// on a large deck - and they used to run between arming
							// the rise and its first frame, which is why the entry
							// animation stuttered. The light reads above stay: the
							// window and the page counter both need them now.
							this._pendingSlidesCollection = slides;
							this.IsShowing = true;
							this.OnShowBegin();
							ConfigHelper.Instance.IsShowing = this.IsShowing;
							SyntacticSugar.CurrentDispatcherBeginInvoke(delegate
							{
								this._stopwatch.Restart();
							});
							ConfigHelper.Instance.IsComOk = true;
							Log.WriteDebug("PPTShowBegin ok: slides=" + this.SlidesCount
								+ " isWps=" + PptVersionManager.IsWps
								+ " path=" + this.PptFilePath);
						}
					}
				}
			}
			catch (Exception ex)
			{
				ConfigHelper.Instance.IsComOk = false;
				// The one place worth a full stack: everything the toolbar needs is
				// decided in here, so a failure means the toolbars will not appear and
				// there is nothing else to look at.
				Log.WriteError("ShowBeginError: " + ex);
			}
		}

		/// <summary>
		/// Reads SlideShowSettings.RangeType, defaulting to "all slides" when it
		/// cannot be read.
		///
		/// The fallback matters: the page counter and the thumbnail list are both
		/// built from RangeType, and guessing wrong would leave the toolbar counting
		/// slides that a custom show does not contain. ppShowAll is the one value
		/// that is always true of what is on screen.
		/// </summary>
		private static PpSlideShowRangeType ReadRangeType(object settings)
		{
			int raw = ComLate.GetInt(settings, "RangeType", (int)PpSlideShowRangeType.ppShowAll);
			try
			{
				return (PpSlideShowRangeType)raw;
			}
			catch (ArgumentException)
			{
				return PpSlideShowRangeType.ppShowAll;
			}
		}

		// Token: 0x06000242 RID: 578 RVA: 0x0000AC08 File Offset: 0x00008E08
		private void OnAppSlideShowEnd(object pres)
		{
			// Click-advance goes back on before anything else, even if we were never
			// showing: leaving it off would leave the user unable to advance normally.
			// msoTrue = 1.
			this.SetSlideAdvanceOnClickProperty(true);
			SyntacticSugar.TryCatchAction(delegate
			{
				if (!this._isShowTypeSpeaker)
				{
					return;
				}
				this.IsShowing = false;
				PptAnalyzer.StopAnalyzeAllAsyn();
				SyntacticSugar.CurrentDispatcherBeginInvoke(delegate
				{
					this._stopwatch.Stop();
				});
				ConfigHelper.Instance.IsShowing = this.IsShowing;
				this._isShowTypeSpeaker = false;
				this.OnShowEnd();
				Log.WriteDebug("PPTShowEnd");
			}, null);
		}

		// Token: 0x06000243 RID: 579 RVA: 0x0000AC23 File Offset: 0x00008E23
		// The window argument is unused: a jump to the next slide from anywhere still
		// means "the current show advanced", and reading the window back off this
		// callback would just re-introduce the typed marshalling this whole class
		// had to give up.
		private void OnSlideShowNextSlide(object wn)
		{
			SyntacticSugar.TryCatchAction(delegate
			{
				if (!this._isShowTypeSpeaker)
				{
					return;
				}
				SyntacticSugar.CurrentDispatcherBeginInvoke(delegate
				{
					this._stopwatch.Restart();
				});
				this.OnNextSlide();
			}, null);
		}

		// Token: 0x0400010C RID: 268
		private int _times;

		// Token: 0x0400010D RID: 269
		// Late-bound, not the Application coclass: casting to it is a QueryInterface
		// for a Microsoft IID, which WPS does not implement - that cast was throwing
		// InvalidCastException and is the reason the whole COM layer went late bound.
		// See ComLate.
		private object _app;

		// Token: 0x0400010E RID: 270
		private Stopwatch _stopwatch;

		// Token: 0x0400010F RID: 271
		private bool _isShowTypeSpeaker;

		// Token: 0x04000110 RID: 272
		// Late-bound like _app: never cast to the typed SlideShowWindow, for the same
		// reason. It is written once, when a show begins, and never nulled - so it
		// must be re-resolved rather than trusted, which is what ComLate does
		// everywhere it is read.
		private object _slideShowWindow;

		// Token: 0x04000111 RID: 273
		private const int PptProcessCheckInterval = 2500;

		// Token: 0x04000112 RID: 274
		private static readonly int MinPageTurningTimerSpan = 30;

		// Token: 0x0400011F RID: 287
		private bool _isCheckPptStateFirstly = true;

		// Token: 0x04000120 RID: 288
		private readonly DispatcherTimer _checkPowerPointTimer = new DispatcherTimer();

		// Token: 0x04000121 RID: 289
		private readonly DispatcherTimer _clearEmptyPowerPointTimer = new DispatcherTimer();

		// Straight ticks of the cleanup timer on which Presentations could not be
		// read at all. Reset by any successful read; at 2.5 s per tick, five of them
		// is 12.5 seconds of a server that answers nothing, which is "gone" rather
		// than "busy".
		private const int UnreadableTicksBeforeCleanup = 5;

		private int _unreadableTicks;
	}
}
