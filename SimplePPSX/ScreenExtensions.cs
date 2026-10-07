using System;
using System.Windows;

namespace SimplePPSX
{
	// Token: 0x02000024 RID: 36
	public static class ScreenExtensions
	{
		// Token: 0x06000132 RID: 306 RVA: 0x000052FA File Offset: 0x000034FA
		public static Point ToWpfPoint(this Point screenPoint)
		{
			return new Point(screenPoint.X / (double)ScreenHelper.DpiX * 96.0, screenPoint.Y / (double)ScreenHelper.DpiY * 96.0);
		}

		// Token: 0x06000133 RID: 307 RVA: 0x00005331 File Offset: 0x00003531
		public static Point ToScreenPoint(this Point wpfPoint)
		{
			return new Point(wpfPoint.X * (double)ScreenHelper.DpiX / 96.0, wpfPoint.Y * (double)ScreenHelper.DpiY / 96.0);
		}
	}
}
