using System;
using System.Diagnostics;

namespace SimplePPSX
{
	// Token: 0x02000017 RID: 23
	public static class ExploreHelper
	{
		// Token: 0x060000D3 RID: 211 RVA: 0x0000438C File Offset: 0x0000258C
		public static void OpenLink(string sUrl)
		{
			try
			{
				Process.Start(sUrl);
			}
			catch (Exception)
			{
				try
				{
					Process.Start(new ProcessStartInfo("IExplore.exe", sUrl));
				}
				catch (Exception)
				{
				}
			}
		}
	}
}
