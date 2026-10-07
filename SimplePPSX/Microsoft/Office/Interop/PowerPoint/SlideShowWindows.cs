using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.PowerPoint
{
	// Token: 0x02000067 RID: 103
	[CompilerGenerated]
	[Guid("91493456-5A91-11CF-8700-00AA0060263B")]
	[TypeIdentifier]
	[ComImport]
	public interface SlideShowWindows : Collection
	{
		// Token: 0x06000379 RID: 889
		void _VtblGap1_2();

		// Token: 0x17000074 RID: 116
		// (get) Token: 0x0600037A RID: 890
		[DispId(11)]
		int Count
		{
			[DispId(11)]
			[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
			get;
		}

		// Token: 0x0600037B RID: 891
		void _VtblGap2_2();

		// Token: 0x17000075 RID: 117
		[DispId(0)]
		SlideShowWindow this[[In] int Index]
		{
			[DispId(0)]
			[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}
	}
}
