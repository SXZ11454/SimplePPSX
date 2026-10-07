using System;
using System.Windows;

namespace SimplePPSX
{
	// Token: 0x0200000F RID: 15
	public static class PointExtensions
	{
		// Token: 0x060000A0 RID: 160 RVA: 0x00003A85 File Offset: 0x00001C85
		public static Point OffsetEx(this Point point, double X, double Y)
		{
			return new Point(point.X + X, point.Y + Y);
		}
	}
}
