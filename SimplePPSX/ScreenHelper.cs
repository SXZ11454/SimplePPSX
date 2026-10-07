using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Media;
using Microsoft.Win32;

namespace SimplePPSX
{
	// Token: 0x02000023 RID: 35
	public class ScreenHelper : INotifyPropertyChanged
	{
		// Token: 0x1700002C RID: 44
		// (get) Token: 0x06000114 RID: 276 RVA: 0x00004E32 File Offset: 0x00003032
		// (set) Token: 0x06000115 RID: 277 RVA: 0x00004E39 File Offset: 0x00003039
		public static int DpiX { get; private set; }

		// Token: 0x1700002D RID: 45
		// (get) Token: 0x06000116 RID: 278 RVA: 0x00004E41 File Offset: 0x00003041
		// (set) Token: 0x06000117 RID: 279 RVA: 0x00004E48 File Offset: 0x00003048
		public static int DpiY { get; private set; }

		// Token: 0x1700002E RID: 46
		// (get) Token: 0x06000118 RID: 280 RVA: 0x00004E50 File Offset: 0x00003050
		// (set) Token: 0x06000119 RID: 281 RVA: 0x00004E57 File Offset: 0x00003057
		public static double DpiFactor { get; private set; }

		// Token: 0x1700002F RID: 47
		// (get) Token: 0x0600011A RID: 282 RVA: 0x00004E5F File Offset: 0x0000305F
		public static Rect MainWindowBounds
		{
			get
			{
				if (global::System.Windows.Application.Current.MainWindow == null)
				{
					return new Rectangle(new global::System.Drawing.Point(0, 0), SystemInformation.PrimaryMonitorSize).ToRect((double)ScreenHelper.DpiX);
				}
				return ScreenHelper.GetWindowBounds(global::System.Windows.Application.Current.MainWindow);
			}
		}

		// Token: 0x17000030 RID: 48
		// (get) Token: 0x0600011B RID: 283 RVA: 0x00004E99 File Offset: 0x00003099
		public static Rect MainWindowActualBounds
		{
			get
			{
				if (global::System.Windows.Application.Current.MainWindow == null)
				{
					return new Rectangle(new global::System.Drawing.Point(0, 0), SystemInformation.PrimaryMonitorSize).ToRect((double)ScreenHelper.DpiX);
				}
				return ScreenHelper.GetWindowActualBounds(global::System.Windows.Application.Current.MainWindow);
			}
		}

		// Token: 0x17000031 RID: 49
		// (get) Token: 0x0600011C RID: 284 RVA: 0x00004ED4 File Offset: 0x000030D4
		public static Rect MainWindowScreenBounds
		{
			get
			{
				Rect screenBoundsFromPoint = ScreenHelper.GetScreenBoundsFromPoint(ScreenHelper.MainWindowActualBounds.Location + new Vector(ScreenHelper.MainWindowActualBounds.Width / 2.0, ScreenHelper.MainWindowActualBounds.Height / 2.0));
				if (ScreenHelper._oldMainWindowScreenBounds == default(Rect))
				{
					ScreenHelper._oldMainWindowScreenBounds = screenBoundsFromPoint;
				}
				else if (ScreenHelper._oldMainWindowScreenBounds.Size != screenBoundsFromPoint.Size)
				{
					ScreenHelper._oldMainWindowScreenBounds = screenBoundsFromPoint;
				}
				return screenBoundsFromPoint;
			}
		}

		// Token: 0x0600011D RID: 285 RVA: 0x00004F68 File Offset: 0x00003168
		public static Rect GetScreenBoundsByWindow(Window window)
		{
			return ScreenHelper.GetScreenBoundsFromPoint(ScreenHelper.GetWindowBounds(window).Location + new Vector(window.Width / 2.0, window.Height / 2.0));
		}

		// Token: 0x17000032 RID: 50
		// (get) Token: 0x0600011E RID: 286 RVA: 0x00004FB2 File Offset: 0x000031B2
		public static int ScreenCount
		{
			get
			{
				return Screen.AllScreens.Length;
			}
		}

		// Token: 0x17000033 RID: 51
		// (get) Token: 0x0600011F RID: 287 RVA: 0x00004FBB File Offset: 0x000031BB
		public static int MainWindowScreenIndex
		{
			get
			{
				return ScreenHelper.ScreenRects.FindIndex((Rect rect) => rect.Contains(ScreenHelper.MainWindowBounds.X + ScreenHelper.MainWindowBounds.Width / 2.0, ScreenHelper.MainWindowBounds.Y + ScreenHelper.MainWindowBounds.Height / 2.0));
			}
		}

		// Token: 0x17000034 RID: 52
		// (get) Token: 0x06000120 RID: 288 RVA: 0x00004FE6 File Offset: 0x000031E6
		public static List<Rect> ScreenRects
		{
			get
			{
				return Screen.AllScreens.Select<Screen, Rect>((Screen screen) => screen.Bounds.ToRect()).ToList<Rect>();
			}
		}

		// Token: 0x06000121 RID: 289 RVA: 0x00005016 File Offset: 0x00003216
		static ScreenHelper()
		{
			SystemEvents.DisplaySettingsChanged += delegate(object sender, EventArgs args)
			{
				ScreenHelper.Update();
				ScreenHelper.OnDisplaySettingsChanged();
			};
			ScreenHelper.Update();
		}

		// Token: 0x06000122 RID: 290 RVA: 0x00005034 File Offset: 0x00003234
		private static void Update()
		{
			PropertyInfo property = typeof(SystemParameters).GetProperty("DpiX", BindingFlags.Static | BindingFlags.NonPublic);
			PropertyInfo property2 = typeof(SystemParameters).GetProperty("Dpi", BindingFlags.Static | BindingFlags.NonPublic);
			ScreenHelper.DpiX = ((property != null) ? ((int)property.GetValue(null, null)) : 96);
			ScreenHelper.DpiY = ((property2 != null) ? ((int)property2.GetValue(null, null)) : 96);
			ScreenHelper.DpiFactor = (double)ScreenHelper.DpiX / 96.0;
		}

		// Token: 0x14000007 RID: 7
		// (add) Token: 0x06000123 RID: 291 RVA: 0x000050C4 File Offset: 0x000032C4
		// (remove) Token: 0x06000124 RID: 292 RVA: 0x000050F8 File Offset: 0x000032F8
		public static event EventHandler DisplaySettingsChanged;

		// Token: 0x06000125 RID: 293 RVA: 0x0000512B File Offset: 0x0000332B
		private static void OnDisplaySettingsChanged()
		{
			EventHandler displaySettingsChanged = ScreenHelper.DisplaySettingsChanged;
			if (displaySettingsChanged == null)
			{
				return;
			}
			displaySettingsChanged(null, EventArgs.Empty);
		}

		// Token: 0x06000126 RID: 294 RVA: 0x00005144 File Offset: 0x00003344
		public static double GetRelativePixel(double pixel)
		{
			Rect mainWindowScreenBounds = ScreenHelper.MainWindowScreenBounds;
			if (mainWindowScreenBounds.Width > mainWindowScreenBounds.Height)
			{
				return pixel * ScreenHelper.MainWindowScreenBounds.Height / 1080.0;
			}
			return pixel * ScreenHelper.MainWindowScreenBounds.Width / 1080.0;
		}

		// Token: 0x06000127 RID: 295 RVA: 0x0000519A File Offset: 0x0000339A
		public static Rect GetScreenBoundsFromPoint(global::System.Windows.Point point)
		{
			return Screen.GetBounds(point.ToWinFormPoint()).ToRect((double)ScreenHelper.DpiX);
		}

		// Token: 0x06000128 RID: 296 RVA: 0x000051B2 File Offset: 0x000033B2
		public static Rect GetWorkingArea(global::System.Windows.Point point)
		{
			return Screen.GetWorkingArea(point.ToWinFormPoint()).ToRect();
		}

		// Token: 0x06000129 RID: 297 RVA: 0x000051C4 File Offset: 0x000033C4
		public static Rect GetWindowBounds(Window window)
		{
			return new Rect(window.Left, window.Top, window.Width, window.Height);
		}

		// Token: 0x0600012A RID: 298 RVA: 0x000051E4 File Offset: 0x000033E4
		public static Rect GetWindowActualBounds(Window window)
		{
			global::System.Windows.Point point = window.PointToScreen(default(global::System.Windows.Point)).ToWpfPoint();
			return new Rect(point.X, point.Y, window.ActualWidth, window.ActualHeight);
		}

		// Token: 0x0600012B RID: 299 RVA: 0x00005228 File Offset: 0x00003428
		public static global::System.Windows.Point GetMousePosition()
		{
			ScreenHelper.POINT point;
			ScreenHelper.GetCursorPos(out point);
			return point;
		}

		// Token: 0x0600012C RID: 300 RVA: 0x00005244 File Offset: 0x00003444
		public static global::System.Windows.Point GetMousePosition(Visual relativeToVisual)
		{
			ScreenHelper.POINT point;
			ScreenHelper.GetCursorPos(out point);
			return relativeToVisual.PointFromScreen(new global::System.Windows.Point((double)point.X, (double)point.Y));
		}

		// Token: 0x0600012D RID: 301
		[DllImport("user32.dll")]
		private static extern bool GetCursorPos(out ScreenHelper.POINT lpPoint);

		// Token: 0x14000008 RID: 8
		// (add) Token: 0x0600012E RID: 302 RVA: 0x00005274 File Offset: 0x00003474
		// (remove) Token: 0x0600012F RID: 303 RVA: 0x000052AC File Offset: 0x000034AC
		public event PropertyChangedEventHandler PropertyChanged;

		// Token: 0x06000130 RID: 304 RVA: 0x000052E1 File Offset: 0x000034E1
		protected virtual void RaisePropertyChanged(string propertyName)
		{
			PropertyChangedEventHandler propertyChanged = this.PropertyChanged;
			if (propertyChanged == null)
			{
				return;
			}
			propertyChanged(this, new PropertyChangedEventArgs(propertyName));
		}

		// Token: 0x0400005F RID: 95
		private static Rect _oldMainWindowScreenBounds;

		// Token: 0x02000091 RID: 145
		private struct POINT
		{
			// Token: 0x06000454 RID: 1108 RVA: 0x0000EFD7 File Offset: 0x0000D1D7
			public static implicit operator global::System.Windows.Point(ScreenHelper.POINT point)
			{
				return new global::System.Windows.Point((double)point.X, (double)point.Y);
			}

			// Token: 0x04000247 RID: 583
			public int X;

			// Token: 0x04000248 RID: 584
			public int Y;
		}
	}
}
