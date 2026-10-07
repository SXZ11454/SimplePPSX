using System;
using System.Runtime.InteropServices;

namespace SimplePPSX
{
	// Token: 0x0200001A RID: 26
	public static class NetworkHelper
	{
		// Token: 0x060000E7 RID: 231
		[DllImport("winInet.dll")]
		private static extern bool InternetGetConnectedState(ref int dwFlag, int dwReserved);

		// Token: 0x17000023 RID: 35
		// (get) Token: 0x060000E8 RID: 232 RVA: 0x000046F8 File Offset: 0x000028F8
		public static bool IsNetworkConnect
		{
			get
			{
				int num = 0;
				return NetworkHelper.InternetGetConnectedState(ref num, 0);
			}
		}
	}
}
