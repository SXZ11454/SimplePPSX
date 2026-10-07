using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.PowerPoint
{
	// Token: 0x02000066 RID: 102
	[CompilerGenerated]
	[Guid("91493453-5A91-11CF-8700-00AA0060263B")]
	[TypeIdentifier]
	[ComImport]
	public interface SlideShowWindow
	{
		// Token: 0x06000372 RID: 882
		void _VtblGap1_2();

		// Token: 0x17000071 RID: 113
		// (get) Token: 0x06000373 RID: 883
		[DispId(2003)]
		SlideShowView View
		{
			[DispId(2003)]
			[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		// Token: 0x17000072 RID: 114
		// (get) Token: 0x06000374 RID: 884
		[DispId(2004)]
		Presentation Presentation
		{
			[DispId(2004)]
			[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		// Token: 0x06000375 RID: 885
		void _VtblGap2_9();

		// Token: 0x17000073 RID: 115
		// (get) Token: 0x06000376 RID: 886
		[DispId(2010)]
		int HWND
		{
			[DispId(2010)]
			[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
			get;
		}

		// Token: 0x06000377 RID: 887
		void _VtblGap3_1();

		// Token: 0x06000378 RID: 888
		[DispId(2012)]
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		void Activate();
	}
}
