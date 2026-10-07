using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.PowerPoint
{
	// Token: 0x02000063 RID: 99
	[CompilerGenerated]
	[Guid("9149345A-5A91-11CF-8700-00AA0060263B")]
	[TypeIdentifier]
	[ComImport]
	public interface SlideShowSettings
	{
		// Token: 0x06000357 RID: 855
		void _VtblGap1_4();

		// Token: 0x17000069 RID: 105
		// (get) Token: 0x06000358 RID: 856
		// (set) Token: 0x06000359 RID: 857
		[DispId(2005)]
		int StartingSlide
		{
			[DispId(2005)]
			[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
			get;
			[DispId(2005)]
			[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
			[param: In]
			set;
		}

		// Token: 0x1700006A RID: 106
		// (get) Token: 0x0600035A RID: 858
		// (set) Token: 0x0600035B RID: 859
		[DispId(2006)]
		int EndingSlide
		{
			[DispId(2006)]
			[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
			get;
			[DispId(2006)]
			[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
			[param: In]
			set;
		}

		// Token: 0x0600035C RID: 860
		void _VtblGap2_5();

		// Token: 0x1700006B RID: 107
		// (get) Token: 0x0600035D RID: 861
		// (set) Token: 0x0600035E RID: 862
		[DispId(2010)]
		PpSlideShowType ShowType
		{
			[DispId(2010)]
			[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
			get;
			[DispId(2010)]
			[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
			[param: In]
			set;
		}

		// Token: 0x0600035F RID: 863
		void _VtblGap3_6();

		// Token: 0x1700006C RID: 108
		// (get) Token: 0x06000360 RID: 864
		// (set) Token: 0x06000361 RID: 865
		[DispId(2014)]
		PpSlideShowRangeType RangeType
		{
			[DispId(2014)]
			[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
			get;
			[DispId(2014)]
			[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
			[param: In]
			set;
		}
	}
}
