using System;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.PowerPoint
{
	// Token: 0x02000069 RID: 105
	[CompilerGenerated]
	[Guid("91493442-5A91-11CF-8700-00AA0060263B")]
	[DefaultMember("Name")]
	[TypeIdentifier]
	[ComImport]
	public interface _Application
	{
		// Token: 0x17000077 RID: 119
		// (get) Token: 0x0600037F RID: 895
		[DispId(2001)]
		Presentations Presentations
		{
			[DispId(2001)]
			[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		// Token: 0x06000380 RID: 896
		void _VtblGap1_3();

		// Token: 0x17000078 RID: 120
		// (get) Token: 0x06000381 RID: 897
		[DispId(2005)]
		Presentation ActivePresentation
		{
			[DispId(2005)]
			[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		// Token: 0x17000079 RID: 121
		// (get) Token: 0x06000382 RID: 898
		[DispId(2006)]
		SlideShowWindows SlideShowWindows
		{
			[DispId(2006)]
			[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		// Token: 0x06000383 RID: 899
		void _VtblGap2_1();

		// Token: 0x1700007A RID: 122
		// (get) Token: 0x06000384 RID: 900
		[DispId(2008)]
		string Path
		{
			[DispId(2008)]
			[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
			[return: MarshalAs(UnmanagedType.BStr)]
			get;
		}

		// Token: 0x1700007B RID: 123
		// (get) Token: 0x06000385 RID: 901
		[DispId(0)]
		string Name
		{
			[DispId(0)]
			[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
			[return: MarshalAs(UnmanagedType.BStr)]
			get;
		}

		// Token: 0x1700007C RID: 124
		// (get) Token: 0x06000386 RID: 902
		// (set) Token: 0x06000387 RID: 903
		[DispId(2009)]
		string Caption
		{
			[DispId(2009)]
			[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
			[return: MarshalAs(UnmanagedType.BStr)]
			get;
			[DispId(2009)]
			[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
			[param: MarshalAs(UnmanagedType.BStr)]
			[param: In]
			set;
		}

		// Token: 0x06000388 RID: 904
		void _VtblGap3_4();

		// Token: 0x1700007D RID: 125
		// (get) Token: 0x06000389 RID: 905
		[DispId(2014)]
		string Version
		{
			[DispId(2014)]
			[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
			[return: MarshalAs(UnmanagedType.BStr)]
			get;
		}

		// Token: 0x0600038A RID: 906
		void _VtblGap4_22();

		// Token: 0x1700007E RID: 126
		// (get) Token: 0x0600038B RID: 907
		[DispId(2031)]
		int HWND
		{
			[DispId(2031)]
			[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
			get;
		}
	}
}
