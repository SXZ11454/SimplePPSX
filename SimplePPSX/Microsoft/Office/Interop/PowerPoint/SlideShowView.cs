using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;

namespace Microsoft.Office.Interop.PowerPoint
{
	// Token: 0x02000065 RID: 101
	[CompilerGenerated]
	[Guid("91493459-5A91-11CF-8700-00AA0060263B")]
	[TypeIdentifier]
	[ComImport]
	public interface SlideShowView
	{
		// Token: 0x06000365 RID: 869
		void _VtblGap1_4();

		// Token: 0x1700006E RID: 110
		// (get) Token: 0x06000366 RID: 870
		// (set) Token: 0x06000367 RID: 871
		[DispId(2005)]
		PpSlideShowPointerType PointerType
		{
			[DispId(2005)]
			[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
			get;
			[DispId(2005)]
			[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
			[param: In]
			set;
		}

		// Token: 0x06000368 RID: 872
		void _VtblGap2_9();

		// Token: 0x1700006F RID: 111
		// (get) Token: 0x06000369 RID: 873
		[DispId(2012)]
		ColorFormat PointerColor
		{
			[DispId(2012)]
			[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		// Token: 0x0600036A RID: 874
		void _VtblGap3_6();

		// Token: 0x0600036B RID: 875
		[DispId(2019)]
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		void Next();

		// Token: 0x0600036C RID: 876
		[DispId(2020)]
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		void Previous();

		// Token: 0x0600036D RID: 877
		[DispId(2021)]
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		void GotoSlide([In] int Index, [In] MsoTriState ResetSlide = MsoTriState.msoTrue);

		// Token: 0x0600036E RID: 878
		void _VtblGap4_3();

		// Token: 0x0600036F RID: 879
		[DispId(2025)]
		[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
		void Exit();

		// Token: 0x06000370 RID: 880
		void _VtblGap5_1();

		// Token: 0x17000070 RID: 112
		// (get) Token: 0x06000371 RID: 881
		[DispId(2027)]
		int CurrentShowPosition
		{
			[DispId(2027)]
			[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
			get;
		}
	}
}
