using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;

namespace Microsoft.Office.Interop.PowerPoint
{
	// Token: 0x02000064 RID: 100
	[CompilerGenerated]
	[Guid("91493471-5A91-11CF-8700-00AA0060263B")]
	[TypeIdentifier]
	[ComImport]
	public interface SlideShowTransition
	{
		// Token: 0x06000362 RID: 866
		void _VtblGap1_2();

		// Token: 0x1700006D RID: 109
		// (get) Token: 0x06000363 RID: 867
		// (set) Token: 0x06000364 RID: 868
		[DispId(2003)]
		MsoTriState AdvanceOnClick
		{
			[DispId(2003)]
			[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
			get;
			[DispId(2003)]
			[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
			[param: In]
			set;
		}
	}
}
