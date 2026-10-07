using System;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.PowerPoint
{
	// Token: 0x0200004F RID: 79
	[CompilerGenerated]
	[DefaultMember("RGB")]
	[Guid("91493452-5A91-11CF-8700-00AA0060263B")]
	[TypeIdentifier]
	[ComImport]
	public interface ColorFormat
	{
		// Token: 0x06000338 RID: 824
		void _VtblGap1_3();

		// Token: 0x17000067 RID: 103
		// (get) Token: 0x06000339 RID: 825
		// (set) Token: 0x0600033A RID: 826
		[DispId(0)]
		int RGB
		{
			[DispId(0)]
			[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
			get;
			[DispId(0)]
			[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
			[param: In]
			set;
		}
	}
}
