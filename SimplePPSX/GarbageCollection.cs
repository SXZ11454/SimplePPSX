using System;

namespace SimplePPSX
{
	// Token: 0x02000015 RID: 21
	public static class GarbageCollection
	{
		// Token: 0x060000C5 RID: 197 RVA: 0x00003FF0 File Offset: 0x000021F0
		public static void Collect()
		{
			GC.Collect();
			GC.WaitForPendingFinalizers();
			GC.Collect();
		}
	}
}
