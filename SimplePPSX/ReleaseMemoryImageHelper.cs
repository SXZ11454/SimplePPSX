using System;
using System.IO;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace SimplePPSX
{
	// Token: 0x0200001E RID: 30
	public static class ReleaseMemoryImageHelper
	{
		// Token: 0x060000F4 RID: 244 RVA: 0x000048DC File Offset: 0x00002ADC
		public static string GetImageFilePath()
		{
			string text = Path.Combine(AppPath.DocumentPath(), ReleaseMemoryImageHelper.ImageName);
			if (File.Exists(text))
			{
				return text;
			}
			byte[] array = new byte[4];
			BitmapSource.Create(1, 1, 96.0, 96.0, PixelFormats.Bgra32, null, array, 4).SaveToFile(text, false);
			return text;
		}

		// Token: 0x0400004D RID: 77
		private static string ImageName = "RleaseMemoryImage.png";
	}
}
