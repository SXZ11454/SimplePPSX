using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Interop.PowerPoint
{
	// Token: 0x02000051 RID: 81
	[CompilerGenerated]
	[ComEventInterface(typeof(EApplication), typeof(EApplication))]
	[TypeIdentifier("91493440-5a91-11cf-8700-00aa0060263b", "Microsoft.Office.Interop.PowerPoint.EApplication_Event")]
	public interface EApplication_Event
	{
		// Token: 0x06000342 RID: 834
		void _VtblGap1_20();

		// Token: 0x1400000C RID: 12
		// (add) Token: 0x06000343 RID: 835
		// (remove) Token: 0x06000344 RID: 836
		event EApplication_SlideShowBeginEventHandler SlideShowBegin;

		// Token: 0x06000345 RID: 837
		void _VtblGap2_2();

		// Token: 0x1400000D RID: 13
		// (add) Token: 0x06000346 RID: 838
		// (remove) Token: 0x06000347 RID: 839
		event EApplication_SlideShowNextSlideEventHandler SlideShowNextSlide;

		// Token: 0x1400000E RID: 14
		// (add) Token: 0x06000348 RID: 840
		// (remove) Token: 0x06000349 RID: 841
		event EApplication_SlideShowEndEventHandler SlideShowEnd;

		// Token: 0x0600034A RID: 842
		void _VtblGap3_32();

		// Token: 0x1400000F RID: 15
		// (add) Token: 0x0600034B RID: 843
		// (remove) Token: 0x0600034C RID: 844
		event EApplication_PresentationCloseFinalEventHandler PresentationCloseFinal;
	}
}
