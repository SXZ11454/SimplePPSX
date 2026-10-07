using System;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.PowerPoint
{
	// Token: 0x02000068 RID: 104
	[CompilerGenerated]
	[DefaultMember("Item")]
	[Guid("91493469-5A91-11CF-8700-00AA0060263B")]
	[TypeIdentifier]
	[ComImport]
	public interface Slides : Collection
	{
		// Token: 0x0600037D RID: 893
		void _VtblGap1_2();

		// Token: 0x17000076 RID: 118
		// (get) Token: 0x0600037E RID: 894
		[DispId(11)]
		int Count
		{
			[DispId(11)]
			[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
			get;
		}
	}
}
