using System;
using System.Windows;

namespace SimplePPSX
{
	// Token: 0x02000019 RID: 25
	public class MouseHelperEventArgs : EventArgs
	{
		// Token: 0x17000022 RID: 34
		// (get) Token: 0x060000E4 RID: 228 RVA: 0x000046BF File Offset: 0x000028BF
		// (set) Token: 0x060000E5 RID: 229 RVA: 0x000046C7 File Offset: 0x000028C7
		public Point MousePoint { get; set; }

		// Token: 0x060000E6 RID: 230 RVA: 0x000046D0 File Offset: 0x000028D0
		public MouseHelperEventArgs(double x, double y)
		{
			x /= ScreenHelper.DpiFactor;
			y /= ScreenHelper.DpiFactor;
			this.MousePoint = new Point(x, y);
		}
	}
}
