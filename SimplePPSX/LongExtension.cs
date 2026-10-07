using System;

namespace SimplePPSX
{
	// Token: 0x02000029 RID: 41
	public static class LongExtension
	{
		// Token: 0x06000160 RID: 352 RVA: 0x00005838 File Offset: 0x00003A38
		public static string ConvertMilliseconds(this long totalMilliseconds)
		{
			if (totalMilliseconds < 100L)
			{
				return "0.1秒以内";
			}
			if (totalMilliseconds >= 100L && totalMilliseconds < 500L)
			{
				return "0.1~0.5秒";
			}
			if (totalMilliseconds >= 500L && totalMilliseconds < 1000L)
			{
				return "0.5~1秒";
			}
			if (totalMilliseconds >= 1000L && totalMilliseconds < 1500L)
			{
				return "1~1.5秒";
			}
			if (totalMilliseconds >= 1500L && totalMilliseconds < 2000L)
			{
				return "1.5~2秒";
			}
			if (totalMilliseconds >= 2000L && totalMilliseconds < 2500L)
			{
				return "2~2.5秒";
			}
			if (totalMilliseconds >= 2500L && totalMilliseconds < 3000L)
			{
				return "2.5~3秒";
			}
			if (totalMilliseconds >= 3000L && totalMilliseconds < 4000L)
			{
				return "3~4秒";
			}
			if (totalMilliseconds >= 4000L && totalMilliseconds < 5000L)
			{
				return "4~5秒";
			}
			if (totalMilliseconds >= 5000L && totalMilliseconds < 10000L)
			{
				return "5~10秒";
			}
			if (totalMilliseconds >= 10000L && totalMilliseconds < 15000L)
			{
				return "10~15秒";
			}
			if (totalMilliseconds >= 15000L && totalMilliseconds < 20000L)
			{
				return "15~20秒";
			}
			if (totalMilliseconds >= 20000L && totalMilliseconds < 25000L)
			{
				return "20~25秒";
			}
			if (totalMilliseconds >= 25000L && totalMilliseconds < 30000L)
			{
				return "20~30秒";
			}
			if (totalMilliseconds >= 30000L && totalMilliseconds < 40000L)
			{
				return "30~40秒";
			}
			if (totalMilliseconds >= 40000L && totalMilliseconds < 50000L)
			{
				return "40~50秒";
			}
			if (totalMilliseconds >= 50000L && totalMilliseconds < 60000L)
			{
				return "50~60秒";
			}
			if (totalMilliseconds >= 60000L && totalMilliseconds < 120000L)
			{
				return "60~120秒";
			}
			if (totalMilliseconds >= 120000L && totalMilliseconds < 180000L)
			{
				return "120~180秒";
			}
			return "180秒以上";
		}
	}
}
