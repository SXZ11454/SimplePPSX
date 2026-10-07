using System;
using System.Drawing;
using System.Windows;

namespace SimplePPSX
{
	// Token: 0x02000025 RID: 37
	internal static class InteropExtension
	{
		// Token: 0x06000134 RID: 308 RVA: 0x00005368 File Offset: 0x00003568
		public static global::System.Drawing.Point ToWinFormPoint(this global::System.Windows.Point point)
		{
			return new global::System.Drawing.Point((int)(point.X + 0.5), (int)(point.Y + 0.5));
		}

		// Token: 0x06000135 RID: 309 RVA: 0x00005393 File Offset: 0x00003593
		public static Rect ToRect(this Rectangle rectangle)
		{
			return new Rect((double)rectangle.X, (double)rectangle.Y, (double)rectangle.Width, (double)rectangle.Height);
		}

		// Token: 0x06000136 RID: 310 RVA: 0x000053BC File Offset: 0x000035BC
		public static Rect ToRect(this Rectangle rectangle, double dpi)
		{
			double num = 96.0 / dpi;
			return new Rect((double)rectangle.X * num, (double)rectangle.Y * num, (double)rectangle.Width * num, (double)rectangle.Height * num);
		}
	}
}
