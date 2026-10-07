using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;

namespace SimplePPSX
{
	// Token: 0x02000041 RID: 65
	public class WindowsApi
	{
		// Token: 0x060002A4 RID: 676
		[DllImport("gdi32.dll")]
		public static extern int GetDeviceCaps(IntPtr hdc, int nIndex);


		// Token: 0x060002A9 RID: 681
		[DllImport("kernel32.dll", SetLastError = true)]
		internal static extern IntPtr CreateFile(string fileName, uint desiredAccess, uint shareMode, uint securityAttributes, uint creationDisposition, uint flagsAndAttributes, int hTemplateFile);

		// Token: 0x060002AA RID: 682
		[DllImport("Kernel32.dll")]
		internal static extern bool ReadFile(IntPtr hFile, byte[] lpBuffer, uint nNumberOfBytesToRead, ref uint lpNumberOfBytesRead, IntPtr lpOverlapped);

		// Token: 0x060002AB RID: 683
		[DllImport("Kernel32.dll")]
		internal static extern bool WriteFile(IntPtr hFile, byte[] lpBuffer, uint nNumberOfBytesToWrite, ref uint lpNumberOfBytesWritten, IntPtr lpOverlapped);

		// Token: 0x060002AC RID: 684
		[DllImport("kernel32.dll")]
		internal static extern bool CloseHandle(IntPtr hObject);

		// Token: 0x060002AD RID: 685
		[DllImport("user32.dll", CharSet = CharSet.Auto)]
		public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

		// Token: 0x060002AE RID: 686
		[DllImport("user32.dll")]
		[return: MarshalAs(UnmanagedType.Bool)]
		internal static extern bool GetWindowRect(HandleRef hWnd, out WindowRect lpRect);

		// Token: 0x060002AF RID: 687
		[DllImport("user32.dll")]
		public static extern IntPtr SetWindowPos(IntPtr hWnd, int hWndInsertAfter, int x, int y, int cx, int cy, int wFlags);

		// Token: 0x060002B0 RID: 688
		[DllImport("user32", EntryPoint = "mouse_event")]
		public static extern int Mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);

		// Token: 0x060002B1 RID: 689
		[DllImport("user32.dll")]
		public static extern bool GetCursorPos(out MousePoint lpPoint);

		// Token: 0x060002B2 RID: 690
		[DllImport("user32.dll")]
		public static extern bool SetCursorPos(int x, int y);

		// Token: 0x060002B3 RID: 691
		[DllImport("user32.dll")]
		public static extern int GetKeyState(int nVirtKey);

		// Token: 0x060002B4 RID: 692
		[DllImport("user32.dll")]
		public static extern bool SetForegroundWindow(IntPtr hWnd);

		// Token: 0x060002B5 RID: 693
		[DllImport("user32.dll")]
		public static extern int IsIconic(IntPtr hwnd);

		// Token: 0x060002B6 RID: 694
		[DllImport("user32.dll")]
		public static extern bool IsWindowVisible(IntPtr hwnd);

		// Token: 0x060002B7 RID: 695
		[DllImport("user32.dll")]
		public static extern int WindowFromPoint(int xPoint, int yPoint);

		// Token: 0x060002B8 RID: 696
		[DllImport("user32.Dll")]
		[return: MarshalAs(UnmanagedType.Bool)]
		public static extern bool EnumChildWindows(IntPtr parentHandle, WindowsApi.EnumWindowProc callback, IntPtr lParam);

		// Token: 0x060002B9 RID: 697
		[DllImport("user32.dll")]
		public static extern int SendMessage(IntPtr hwnd, int wMsg, int wParam, int lParam);

		// Token: 0x060002BA RID: 698
		[DllImport("user32.dll")]
		public static extern bool PostMessage(IntPtr hwnd, int msg, uint wParam, uint lParam);

		// Token: 0x060002BB RID: 699
		[DllImport("user32")]
		public static extern uint SetWindowLong(IntPtr hwnd, int nIndex, uint dwNewLong);

		// Token: 0x060002BC RID: 700
		[DllImport("user32")]
		public static extern uint GetWindowLong(IntPtr hwnd, int nIndex);

		// Token: 0x060002BD RID: 701
		[DllImport("gdi32.dll")]
		public static extern bool DeleteDC([In] IntPtr hdc);

		// Token: 0x060002BE RID: 702
		[DllImport("gdi32.dll")]
		public static extern bool BitBlt(IntPtr hdcDest, int xDest, int yDest, int wDest, int hDest, IntPtr hdcSource, int xSrc, int ySrc, CopyPixelOperation rop);

		// Token: 0x060002BF RID: 703
		[DllImport("gdi32.dll", SetLastError = true)]
		public static extern IntPtr SelectObject([In] IntPtr hdc, [In] IntPtr hgdiobj);

		// Token: 0x060002C0 RID: 704
		[DllImport("gdi32.dll")]
		public static extern IntPtr CreateCompatibleBitmap([In] IntPtr hdc, int nWidth, int nHeight);

		// Token: 0x060002C1 RID: 705
		[DllImport("gdi32.dll", SetLastError = true)]
		public static extern IntPtr CreateCompatibleDC([In] IntPtr hdc);

		// Token: 0x060002C2 RID: 706
		[DllImport("user32.dll")]
		public static extern IntPtr GetWindowDC(IntPtr hWnd);

		// Token: 0x060002C3 RID: 707
		[DllImport("User32.dll")]
		public static extern int SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

		// Token: 0x060002C4 RID: 708
		[DllImport("user32.dll")]
		public static extern int SetWindowLong(IntPtr window, int index, int value);

		// Token: 0x060002C5 RID: 709
		[DllImport("user32.dll", CharSet = CharSet.Auto)]
		public static extern bool GetCursorPos(out global::System.Windows.Point pt);

		// Token: 0x060002C6 RID: 710
		[DllImport("User32.dll")]
		public static extern int FindWindow(string lpClassName, string lpWindowName);

		// Token: 0x060002C7 RID: 711
		[DllImport("user32.dll")]
		public static extern bool ReleaseDC(IntPtr hWnd, IntPtr hDc);

		// Token: 0x060002C8 RID: 712
		[DllImport("user32.dll")]
		public static extern bool ReleaseCapture();

		// Token: 0x060002C9 RID: 713
		[DllImport("user32.dll")]
		public static extern IntPtr GetDesktopWindow();

		// Token: 0x060002CA RID: 714
		[DllImport("user32.dll")]
		public static extern IntPtr WindowFromPoint(MousePoint point);

		// Token: 0x060002CB RID: 715
		[DllImport("gdi32.dll")]
		public static extern bool DeleteObject(IntPtr ptr);

		// Token: 0x060002CC RID: 716
		[DllImport("user32.dll", SetLastError = true)]
		public static extern IntPtr GetWindow(IntPtr hWnd, GetWindowCmd uCmd);

		// Token: 0x060002CD RID: 717
		[DllImport("user32.dll", EntryPoint = "SystemParametersInfo", SetLastError = true)]
		public static extern bool SystemParametersInfoSet(uint action, uint uiParam, uint vparam, uint init);

		// Token: 0x060002CE RID: 718
		[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

		// Needed by FindWindowByClassName, which locates the slideshow window because
		// its HWND is not reachable through COM (see that method for the reasoning).
		[DllImport("user32.dll")]
		public static extern bool EnumWindows(EnumWindowProc callback, IntPtr lParam);

		[DllImport("user32.dll", CharSet = CharSet.Unicode)]
		public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

		// Token: 0x060002D4 RID: 724 RVA: 0x0000BD89 File Offset: 0x00009F89
		internal static IntPtr CreateDeviceFile(string device)
		{
			return WindowsApi.CreateFile(device, 3221225472U, 0U, 0U, 3U, 1073741824U, 0);
		}

		// Token: 0x060002D5 RID: 725 RVA: 0x0000BD9F File Offset: 0x00009F9F
		internal static void SetPopupParameter()
		{
			WindowsApi.SystemParametersInfoSet(28U, 0U, 0U, 0U);
		}

		/// <summary>
		/// Returns the first visible top-level window whose class name matches, or
		/// IntPtr.Zero if there is none.
		///
		/// Added because the slideshow's HWND cannot be read through COM: it is on
		/// SlideShowWindow's vtable but not on its dispatch interface, so an
		/// IDispatch call for it fails with DISP_E_MEMBERNOTFOUND. See
		/// PptOperator.ResolveSlideShowWindowHandle.
		///
		/// Only top-level windows are considered, and only visible ones: the slideshow
		/// is a top-level window, while the editing window and all of PowerPoint's
		/// helper windows are either invisible or not the class asked for.
		/// </summary>
		internal static IntPtr FindWindowByClassName(string className)
		{
			if (string.IsNullOrEmpty(className))
			{
				return IntPtr.Zero;
			}

			IntPtr found = IntPtr.Zero;
			WindowsApi.EnumWindowProc callback = delegate(IntPtr handle, IntPtr pointer)
			{
				if (!WindowsApi.IsWindowVisible(handle))
				{
					return true;
				}

				StringBuilder classBuffer = new StringBuilder(256);
				WindowsApi.GetClassName(handle, classBuffer, classBuffer.Capacity);
				if (classBuffer.ToString() == className)
				{
					found = handle;
					return false;
				}
				return true;
			};

			WindowsApi.EnumWindows(callback, IntPtr.Zero);
			return found;
		}

		// Token: 0x060002D6 RID: 726 RVA: 0x0000BDAC File Offset: 0x00009FAC
		public static List<IntPtr> GetChildWindows(IntPtr parent)
		{
			List<IntPtr> list = new List<IntPtr>();
			GCHandle gchandle = GCHandle.Alloc(list);
			try
			{
				WindowsApi.EnumWindowProc enumWindowProc = new WindowsApi.EnumWindowProc(WindowsApi.EnumWindow);
				WindowsApi.EnumChildWindows(parent, enumWindowProc, GCHandle.ToIntPtr(gchandle));
			}
			finally
			{
				if (gchandle.IsAllocated)
				{
					gchandle.Free();
				}
			}
			list.Add(parent);
			return list;
		}

		// Token: 0x060002D7 RID: 727 RVA: 0x0000BE0C File Offset: 0x0000A00C
		private static bool EnumWindow(IntPtr handle, IntPtr pointer)
		{
			List<IntPtr> list = GCHandle.FromIntPtr(pointer).Target as List<IntPtr>;
			if (list == null)
			{
				throw new InvalidCastException("GCHandle Target could not be cast as List<IntPtr>");
			}
			list.Add(handle);
			return true;
		}


		// Token: 0x020000A7 RID: 167
		// (Invoke) Token: 0x0600049A RID: 1178
		public delegate bool EnumWindowProc(IntPtr hWnd, IntPtr parameter);


		// Token: 0x020000A9 RID: 169
		public class Flagsandattributes
		{
			// Token: 0x04000284 RID: 644
			public const uint FileFlagWriteThrough = 2147483648U;

			// Token: 0x04000285 RID: 645
			public const uint FileFlagOverlapped = 1073741824U;

			// Token: 0x04000286 RID: 646
			public const uint FileFlagNoBuffering = 536870912U;

			// Token: 0x04000287 RID: 647
			public const uint FileFlagRandomAccess = 268435456U;

			// Token: 0x04000288 RID: 648
			public const uint FileFlagSequentialScan = 134217728U;

			// Token: 0x04000289 RID: 649
			public const uint FileFlagDeleteOnClose = 67108864U;

			// Token: 0x0400028A RID: 650
			public const uint FileFlagBackupSemantics = 33554432U;

			// Token: 0x0400028B RID: 651
			public const uint FileFlagPosixSemantics = 16777216U;

			// Token: 0x0400028C RID: 652
			public const uint FileFlagOpenReparsePoint = 2097152U;

			// Token: 0x0400028D RID: 653
			public const uint FileFlagOpenNoRecall = 1048576U;

			// Token: 0x0400028E RID: 654
			public const uint FileFlagFirstPipeInstance = 524288U;
		}

		// Token: 0x020000AA RID: 170
		public class SetWindowPosFlags
		{
			// Token: 0x0400028F RID: 655
			public const short SwpNomove = 2;

			// Token: 0x04000290 RID: 656
			public const short SwpNosize = 1;

			// Token: 0x04000291 RID: 657
			public const short SwpNozorder = 4;

			// Token: 0x04000292 RID: 658
			public const int SwpShowwindow = 64;
		}

		// Token: 0x020000AB RID: 171
		public class MouseFlags
		{
			// Token: 0x04000293 RID: 659
			public const int VkLbutton = 1;

			// Token: 0x04000294 RID: 660
			public const int MouseeventfLeftdown = 2;

			// Token: 0x04000295 RID: 661
			public const int MouseeventfLeftup = 4;

			// Token: 0x04000296 RID: 662
			public const int MouseeventfAbsolute = 32768;
		}


		// Token: 0x020000AE RID: 174
		public class Creationdisposition
		{
			// Token: 0x0400029D RID: 669
			public const uint CreateNew = 1U;

			// Token: 0x0400029E RID: 670
			public const uint CreateAlways = 2U;

			// Token: 0x0400029F RID: 671
			public const uint OpenExisting = 3U;

			// Token: 0x040002A0 RID: 672
			public const uint OpenAlways = 4U;

			// Token: 0x040002A1 RID: 673
			public const uint TruncateExisting = 5U;
		}

		// Token: 0x020000AF RID: 175
		public class Desiredaccess
		{
			// Token: 0x040002A2 RID: 674
			public const uint GenericRead = 2147483648U;

			// Token: 0x040002A3 RID: 675
			public const uint GenericWrite = 1073741824U;

			// Token: 0x040002A4 RID: 676
			public const uint GenericExecute = 536870912U;

			// Token: 0x040002A5 RID: 677
			public const uint GenericAll = 268435456U;
		}

		// Token: 0x020000B0 RID: 176
		public class WindowFlags
		{
			// Token: 0x040002A6 RID: 678
			public const uint WsExLayered = 524288U;

			// Token: 0x040002A7 RID: 679
			public const int WsExTransparent = 32;

			// Token: 0x040002A8 RID: 680
			public const int GwlExstyle = -20;
		}
	}
}
