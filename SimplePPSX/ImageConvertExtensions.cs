using System;
using System.IO;
using System.Windows.Media.Imaging;

namespace SimplePPSX
{
	/// <summary>
	/// Saving a WPF bitmap to a file.
	///
	/// The old signature took a System.Drawing.Imaging.ImageFormat, which only
	/// exists in System.Drawing.Common. Everything the app writes is a PNG, so the
	/// format argument is gone and the encoder is chosen here instead - which also
	/// keeps this file free of the System.Drawing dependency.
	/// </summary>
	public static class ImageConvertExtensions
	{
		public static void SaveToFile(this BitmapSource bitmap, string fileName, bool overwrite = false)
		{
			if (bitmap == null)
			{
				throw new ArgumentNullException("bitmap");
			}
			if (File.Exists(fileName) && !overwrite)
			{
				throw new IOException("File " + fileName + " already Exists.");
			}
			BitmapEncoder bitmapEncoder = new PngBitmapEncoder();
			bitmapEncoder.Frames.Add(BitmapFrame.Create(bitmap));
			using (Stream stream = File.Create(fileName))
			{
				bitmapEncoder.Save(stream);
			}
		}
	}
}
