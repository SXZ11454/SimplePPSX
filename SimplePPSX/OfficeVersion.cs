using System;

namespace SimplePPSX
{
	// Token: 0x02000008 RID: 8
	internal enum OfficeVersion
	{
		// Token: 0x0400003D RID: 61
		Unknown,
		// Token: 0x0400003E RID: 62
		Office2007,
		// Token: 0x0400003F RID: 63
		Office2010,
		// Token: 0x04000040 RID: 64
		Office2013,
		// Token: 0x04000041 RID: 65
		Office2016,
		// Token: 0x04000042 RID: 66
		Office2019,
		/// <summary>WPS Presentation (wpp.exe). Not versioned like Office.</summary>
		WPS
	}
}
