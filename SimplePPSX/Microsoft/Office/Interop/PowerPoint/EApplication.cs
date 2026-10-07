using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.PowerPoint
{
	// Token: 0x02000050 RID: 80
	[CompilerGenerated]
	[Guid("914934C2-5A91-11CF-8700-00AA0060263B")]
	[TypeIdentifier]
	[ComImport]
	public interface EApplication
	{
		// Token: 0x0600033B RID: 827
		void _VtblGap1_10();

		// Token: 0x0600033C RID: 828
		[DispId(2011)]
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		void SlideShowBegin([MarshalAs(UnmanagedType.Interface)] [In] SlideShowWindow Wn);

		// Token: 0x0600033D RID: 829
		void _VtblGap2_1();

		// Token: 0x0600033E RID: 830
		[DispId(2013)]
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		void SlideShowNextSlide([MarshalAs(UnmanagedType.Interface)] [In] SlideShowWindow Wn);

		// Token: 0x0600033F RID: 831
		[DispId(2014)]
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		void SlideShowEnd([MarshalAs(UnmanagedType.Interface)] [In] Presentation Pres);

		// Token: 0x06000340 RID: 832
		void _VtblGap3_16();

		// Token: 0x06000341 RID: 833
		[DispId(2031)]
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		void PresentationCloseFinal([MarshalAs(UnmanagedType.Interface)] [In] Presentation Pres);
	}
}
