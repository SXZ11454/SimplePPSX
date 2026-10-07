using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.PowerPoint
{
	// Token: 0x0200006B RID: 107
	[CompilerGenerated]
	[Guid("9149349D-5A91-11CF-8700-00AA0060263B")]
	[TypeIdentifier]
	[ComImport]
	public interface _Presentation
	{
		// Token: 0x0600038F RID: 911
		void _VtblGap1_10();

		// Token: 0x17000081 RID: 129
		// (get) Token: 0x06000390 RID: 912
		[DispId(2011)]
		Slides Slides
		{
			[DispId(2011)]
			[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		// Token: 0x06000391 RID: 913
		void _VtblGap2_3();

		// Token: 0x17000082 RID: 130
		// (get) Token: 0x06000392 RID: 914
		[DispId(2015)]
		SlideShowSettings SlideShowSettings
		{
			[DispId(2015)]
			[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		// Token: 0x06000393 RID: 915
		void _VtblGap3_8();

		// Token: 0x17000083 RID: 131
		// (get) Token: 0x06000394 RID: 916
		[DispId(2024)]
		string FullName
		{
			[DispId(2024)]
			[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
			[return: MarshalAs(UnmanagedType.BStr)]
			get;
		}
	}
}
