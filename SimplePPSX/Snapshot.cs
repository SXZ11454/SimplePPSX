using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media.Imaging;

namespace SimplePPSX
{
	// Token: 0x0200002B RID: 43
	public static class Snapshot
	{
		// Token: 0x06000167 RID: 359 RVA: 0x00005D48 File Offset: 0x00003F48
		public static BitmapSource Take(double x, double y, double width, double height)
		{
			IntPtr desktopWindow = WindowsApi.GetDesktopWindow();
			IntPtr windowDC = WindowsApi.GetWindowDC(desktopWindow);
			IntPtr intPtr = WindowsApi.CreateCompatibleDC(windowDC);
			IntPtr intPtr2 = WindowsApi.CreateCompatibleBitmap(windowDC, (int)width, (int)height);
			IntPtr intPtr3 = WindowsApi.SelectObject(intPtr, intPtr2);
			WindowsApi.BitBlt(intPtr, 0, 0, (int)width, (int)height, windowDC, (int)x, (int)y, (CopyPixelOperation)1087111200);
			Bitmap bitmap = Image.FromHbitmap(intPtr2);
			WindowsApi.SelectObject(intPtr, intPtr3);
			WindowsApi.DeleteObject(intPtr2);
			WindowsApi.DeleteDC(intPtr);
			WindowsApi.ReleaseDC(desktopWindow, windowDC);
			IntPtr hbitmap = bitmap.GetHbitmap();
			BitmapSource bitmapSource = Imaging.CreateBitmapSourceFromHBitmap(hbitmap, IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
			if (!WindowsApi.DeleteObject(hbitmap))
			{
				throw new Win32Exception("gdi32.dll error occurred when deleting Hbitmap object.");
			}
			return bitmapSource;
		}

		// Token: 0x06000168 RID: 360 RVA: 0x00005DE8 File Offset: 0x00003FE8
		public static BitmapSource Take(Rect rect)
		{
			if (Math.Abs(rect.Width) < 1.0 || Math.Abs(rect.Height) < 1.0)
			{
				return null;
			}
			return Snapshot.Take(rect.X, rect.Y, rect.Width, rect.Height);
		}

		// Token: 0x06000169 RID: 361 RVA: 0x00005E46 File Offset: 0x00004046
		public static BitmapSource Take(Rectangle rectangle)
		{
			return Snapshot.Take(new Rect((double)rectangle.X, (double)rectangle.Y, (double)rectangle.Width, (double)rectangle.Height));
		}
	}
}
