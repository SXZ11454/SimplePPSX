using System;
using System.Windows;

namespace SimplePPSX
{
	// Token: 0x02000047 RID: 71
	public struct MousePoint
	{
		// Token: 0x060002DD RID: 733 RVA: 0x0000BE68 File Offset: 0x0000A068
		public MousePoint(double x, double y)
		{
			this.X = (int)x;
			this.Y = (int)y;
		}

		// Token: 0x060002DE RID: 734 RVA: 0x0000BE7A File Offset: 0x0000A07A
		public static implicit operator Point(MousePoint mousePoint)
		{
			return new Point((double)mousePoint.X, (double)mousePoint.Y);
		}

		// Token: 0x04000172 RID: 370
		public int X;

		// Token: 0x04000173 RID: 371
		public int Y;
	}
}
