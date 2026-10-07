using System;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.PowerPoint
{
	// Token: 0x0200005F RID: 95
	[CompilerGenerated]
	[Guid("91493462-5A91-11CF-8700-00AA0060263B")]
	[DefaultMember("Item")]
	[TypeIdentifier]
	[ComImport]
	public interface Presentations : Collection
	{
		// Token: 0x06000355 RID: 853
		void _VtblGap1_2();

		// Token: 0x17000068 RID: 104
		// (get) Token: 0x06000356 RID: 854
		[DispId(11)]
		int Count
		{
			[DispId(11)]
			[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
			get;
		}
	}
}
