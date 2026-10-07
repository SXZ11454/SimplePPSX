using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.PowerPoint
{
	// Token: 0x0200006C RID: 108
	[CompilerGenerated]
	[Guid("9149346A-5A91-11CF-8700-00AA0060263B")]
	[TypeIdentifier]
	[ComImport]
	public interface _Slide
	{
		// Token: 0x06000395 RID: 917
		void _VtblGap1_4();

		// Token: 0x17000084 RID: 132
		// (get) Token: 0x06000396 RID: 918
		[DispId(2005)]
		SlideShowTransition SlideShowTransition
		{
			[DispId(2005)]
			[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		// Token: 0x06000397 RID: 919
		void _VtblGap2_15();

		// Token: 0x17000085 RID: 133
		// (get) Token: 0x06000398 RID: 920
		[DispId(2018)]
		int SlideIndex
		{
			[DispId(2018)]
			[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
			get;
		}

		// Token: 0x06000399 RID: 921
		void _VtblGap3_6();

		// Token: 0x17000086 RID: 134
		// (get) Token: 0x0600039A RID: 922
		[DispId(2023)]
		Master Master
		{
			[DispId(2023)]
			[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		// Token: 0x0600039B RID: 923
		void _VtblGap4_1();

		// Token: 0x0600039C RID: 924
		[DispId(2025)]
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		void Export([MarshalAs(UnmanagedType.BStr)] [In] string FileName, [MarshalAs(UnmanagedType.BStr)] [In] string FilterName, [In] int ScaleWidth = 0, [In] int ScaleHeight = 0);
	}
}
