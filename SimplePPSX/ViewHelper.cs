using System;

namespace SimplePPSX
{
	// Token: 0x0200001F RID: 31
	public class ViewHelper
	{
		// Token: 0x17000025 RID: 37
		// (get) Token: 0x060000F8 RID: 248 RVA: 0x0000494E File Offset: 0x00002B4E
		// (set) Token: 0x060000F9 RID: 249 RVA: 0x00004955 File Offset: 0x00002B55
		public static LeftArea LeftArea { get; private set; }

		// Token: 0x17000026 RID: 38
		// (get) Token: 0x060000FA RID: 250 RVA: 0x0000495D File Offset: 0x00002B5D
		// (set) Token: 0x060000FB RID: 251 RVA: 0x00004964 File Offset: 0x00002B64
		public static RightArea RightArea { get; private set; }

		// Token: 0x17000027 RID: 39
		// (get) Token: 0x060000FC RID: 252 RVA: 0x0000496C File Offset: 0x00002B6C
		// (set) Token: 0x060000FD RID: 253 RVA: 0x00004973 File Offset: 0x00002B73
		public static BottomArea BottomArea { get; private set; }

		// Token: 0x17000028 RID: 40
		// (get) Token: 0x060000FE RID: 254 RVA: 0x0000497B File Offset: 0x00002B7B
		// (set) Token: 0x060000FF RID: 255 RVA: 0x00004982 File Offset: 0x00002B82
		public static LeftBottomArea LeftBottomArea { get; private set; }

		// Token: 0x17000029 RID: 41
		// (get) Token: 0x06000100 RID: 256 RVA: 0x0000498A File Offset: 0x00002B8A
		// (set) Token: 0x06000101 RID: 257 RVA: 0x00004991 File Offset: 0x00002B91
		public static RightBottomArea RightBottomArea { get; private set; }

		// Token: 0x1700002A RID: 42
		// (get) Token: 0x06000102 RID: 258 RVA: 0x00004999 File Offset: 0x00002B99
		// (set) Token: 0x06000103 RID: 259 RVA: 0x000049A0 File Offset: 0x00002BA0
		public static MainWindow MainWindow { get; private set; }

		// The magnifier and the embedded whiteboard were removed from the toolbar and
		// deleted outright, so ViewHelper no longer has a slot for either of them.

		// Token: 0x06000105 RID: 261 RVA: 0x000049DC File Offset: 0x00002BDC
		public void Initialize()
		{
			PptOperator.Instance.ShowBegin += this.Instance_ShowBegin;
			PptOperator.Instance.PowerPointAttached += this.Instance_PowerPointAttached;

			if (ViewHelper.IsPreloadEnabled)
			{
				// Preload: build the whole UI at start-up and present the window while
				// it is still showing nothing - it starts at opacity 0 with every
				// toolbar host collapsed. All the first presentation has to do then is
				// move the toolbars onto the screen, which is what the entry animation
				// needs: a window that has never been presented takes longer than the
				// whole 0.5 second rise to come up, and the rise is over before a
				// single frame of it is drawn.
				ViewHelper.PreloadOverlay("start-up");
			}
			else if (ProcessHelper.HasPowerPointProcess())
			{
				// The switch is off, but PowerPoint was already open when this app
				// started - so the attach and its PowerPointAttached event both
				// happened before this method subscribed. The window still has to be
				// up before the first presentation, so preload now instead.
				ViewHelper.PreloadOverlay("powerpoint already open");
			}
		}

		/// <summary>
		/// Build and present the overlay window before the first presentation.
		/// Defaults to on: without it the window is presented for the first time in
		/// the middle of the show-begin path, which is what the entry animation cannot
		/// afford. See <see cref="PreloadOverlay"/>.
		/// </summary>
		public static bool IsPreloadEnabled
		{
			get
			{
				return ConfigHelper.IsPreload;
			}
			set
			{
				ConfigHelper.IsPreload = value;
			}
		}

		/// <summary>
		/// Fetch the whole deck's thumbnails as soon as a presentation starts,
		/// rather than only when the thumbnail panel is first opened. On by
		/// default - the list is complete the first time the panel opens, at the
		/// cost of a background pass nobody is watching yet. Off defers that pass
		/// to the first open; see ConfigController.IsPrefetch.
		/// </summary>
		public static bool IsPrefetchEnabled
		{
			get
			{
				return ConfigHelper.IsPrefetch;
			}
			set
			{
				ConfigHelper.IsPrefetch = value;
			}
		}

		/// <summary>
		/// The stored colour mode - system / light / dark. The same chain as
		/// every other switch; applying it to ModernWpf and to the toolbar
		/// brushes is ToolbarTheme's job.
		/// </summary>
		public static string ThemeMode
		{
			get
			{
				return ConfigHelper.ThemeMode;
			}
			set
			{
				ConfigHelper.ThemeMode = value;
			}
		}

		/// <summary>
		/// How wide thumbnails are decoded, in device-independent pixels; 0
		/// decodes at the export's own size. The same chain as every other
		/// setting - chosen in the settings page, read by MainWindow's thumbnail
		/// loader each time an image is decoded, so a change takes effect on the
		/// next presentation without a restart.
		/// </summary>
		public static int ThumbnailDecodeWidth
		{
			get
			{
				return ConfigHelper.ThumbnailDecodeWidth;
			}
			set
			{
				ConfigHelper.ThumbnailDecodeWidth = value;
			}
		}

		/// <summary>
		/// The toolbar backdrop opacity of the default (hover) state, in whole
		/// percent. Same chain as every other setting: chosen in the settings
		/// page, pushed onto ToolbarAppearance once at start-up so a restart
		/// keeps the value the user last picked.
		/// </summary>
		public static int ToolbarOpacityPercent
		{
			get
			{
				return ConfigHelper.ToolbarOpacityPercent;
			}
			set
			{
				ConfigHelper.ToolbarOpacityPercent = value;
			}
		}

		/// <summary>
		/// The toolbar backdrop opacity of the idle ("not clicked") state, in
		/// whole percent.
		/// </summary>
		public static int ToolbarIdleOpacityPercent
		{
			get
			{
				return ConfigHelper.ToolbarIdleOpacityPercent;
			}
			set
			{
				ConfigHelper.ToolbarIdleOpacityPercent = value;
			}
		}

		/// <summary>
		/// Builds the overlay and shows its window, leaving it invisible.
		///
		/// The window comes up at opacity 0 with every toolbar host collapsed, so
		/// presenting it here shows nothing on screen and buys the first presentation
		/// a window that is already laid out, already styled and already presenting.
		///
		/// Idempotent, because three moments call it: start-up (preload switch on),
		/// the moment PowerPoint is opened (switch off), and the case where
		/// PowerPoint was already open before this app started.
		/// </summary>
		internal static void PreloadOverlay(string reason)
		{
			if (ViewHelper.MainWindow != null && ViewHelper.MainWindow.IsLoaded)
			{
				return;
			}

			// A preload that fails must cost only the preload: it is on by default,
			// and an exception escaping here would take the whole start-up down with
			// it. The show-begin path builds everything anyway.
			SyntacticSugar.TryCatchAction(delegate
			{
				// Post-or-inline: this can run on PowerPoint's event thread, which
				// must never wait on the UI thread (see
				// SyntacticSugar.DispatcherPostOrInline). The inner try keeps a
				// failed preload costing only the preload even when posted - once
				// posted, the outer try can no longer see the exception.
				SyntacticSugar.DispatcherPostOrInline(delegate
				{
					SyntacticSugar.TryCatchAction(delegate
					{
						ViewHelper.BuildViews();
						if (ViewHelper.MainWindow == null || ViewHelper.MainWindow.IsLoaded)
						{
							return;
						}
						ViewHelper.MainWindow.Show();
						Log.WriteDebug("Overlay window presented early (" + reason + ")");
					}, null);
				});
			}, null);
		}

		// Token: 0x06000106 RID: 262 RVA: 0x000049F4 File Offset: 0x00002BF4
		private void Instance_ShowBegin(object sender, EventArgs e)
		{
			// The whole handler is posted to the UI thread rather than run on
			// PowerPoint's event thread: an event callback must go back to
			// PowerPoint's pump immediately (see
			// SyntacticSugar.DispatcherPostOrInline for the crossing that used to
			// stall both sides). The alreadySubscribed test lives on the UI thread
			// too, since MainWindow's subscription state belongs to it - and it is
			// still evaluated BEFORE BuildViews, which is what subscribes the
			// window.
			//
			// A window that already exists subscribes ShowBegin for itself (in its
			// constructor), so the event is already on its way to it - calling it here
			// as well would run the whole show-begin path twice: two entry animations
			// armed and two analysis passes queued. Only the first, not-yet-built
			// window needs the call.
			SyntacticSugar.DispatcherPostOrInline(delegate
			{
				SyntacticSugar.TryCatchAction(delegate
				{
					bool alreadySubscribed = ViewHelper.MainWindow != null && ViewHelper.MainWindow.IsShowBeginSubscribed;
					ViewHelper.BuildViews();
					if (!alreadySubscribed && ViewHelper.MainWindow != null)
					{
						ViewHelper.MainWindow._pptOperatorShowBegin(null, null);
					}
				}, null);
			});
			PptOperator.Instance.ShowBegin -= this.Instance_ShowBegin;
		}

		/// <summary>
		/// PowerPoint has been opened and attached: with the preload switch off, this
		/// is where the window gets built and presented, so that it is loaded by the
		/// time the first presentation starts.
		/// </summary>
		private void Instance_PowerPointAttached(object sender, EventArgs e)
		{
			if (ViewHelper.IsPreloadEnabled)
			{
				// Already done at start-up; PreloadOverlay would return early anyway.
				return;
			}
			ViewHelper.PreloadOverlay("powerpoint opened");
		}

		/// <summary>构建全部工具面板与主窗口（幂等，已构建则直接返回）。</summary>
		private static void BuildViews()
		{
			if (ViewHelper.MainWindow != null)
			{
				return;
			}
			ViewHelper.LeftArea = new LeftArea();
			ViewHelper.RightArea = new RightArea();
			ViewHelper.LeftBottomArea = new LeftBottomArea();
			ViewHelper.RightBottomArea = new RightBottomArea();
			ViewHelper.BottomArea = new BottomArea();
			ViewHelper.MainWindow = new MainWindow();
			ControlHelper.SetAltF4Enabled(ViewHelper.MainWindow, false);
		}

		// Token: 0x06000107 RID: 263 RVA: 0x00004A47 File Offset: 0x00002C47
		public static void StopMainWindowUpdateTimerAndMouseGesture()
		{
			ViewHelper.MainWindow.IsNeedUpdateTimer = false;
			ViewHelper.MainWindow.UpdateTimer.Stop();
			MouseHelper.Instance.Stop();
		}

		// Token: 0x06000108 RID: 264 RVA: 0x00004A6D File Offset: 0x00002C6D
		public static void BeginMainWindowUpdateTimerAndMouseGesture()
		{
			ViewHelper.MainWindow.IsNeedUpdateTimer = true;
			ViewHelper.MainWindow.UpdateTimer.Start();
			MouseHelper.Instance.Start();
		}
	}
}
