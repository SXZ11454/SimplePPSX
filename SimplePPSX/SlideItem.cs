using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace SimplePPSX
{
	/// <summary>
	/// One thumbnail entry. The class is a UserControl only because the old code
	/// needed a DependencyObject to hang the properties on; it is used as a plain
	/// data item by the ListBox in SlideThumbnailPanel.
	/// </summary>
	public class SlideItem : UserControl
	{
		/// <summary>
		/// The thumbnail itself. The type is ImageSource rather than BitmapImage so it can
		/// also carry a WriteableBitmap, not just a decoded file - only the wider type
		/// serves both.
		/// </summary>
		public ImageSource BitmapImage
		{
			get
			{
				return (ImageSource)base.GetValue(SlideItem.BitmapImageProperty);
			}
			set
			{
				base.SetValue(SlideItem.BitmapImageProperty, value);
			}
		}

		/// <summary>Slide number.</summary>
		public int Index
		{
			get
			{
				return (int)base.GetValue(SlideItem.IndexProperty);
			}
			set
			{
				base.SetValue(SlideItem.IndexProperty, value);
			}
		}

		/// <summary>"True" once a thumbnail image is available, drives the placeholder.</summary>
		public string IsPptThumbnail
		{
			get
			{
				return (string)base.GetValue(SlideItem.IsPptThumbnailProperty);
			}
			set
			{
				base.SetValue(SlideItem.IsPptThumbnailProperty, value);
			}
		}

		public static readonly DependencyProperty IsPptThumbnailProperty = DependencyProperty.Register("IsPptThumbnail", typeof(string), typeof(SlideItem), new PropertyMetadata(null));

		public static readonly DependencyProperty IndexProperty = DependencyProperty.Register("Index", typeof(int), typeof(SlideItem), new PropertyMetadata(0));

		public static readonly DependencyProperty BitmapImageProperty = DependencyProperty.Register("BitmapImage", typeof(ImageSource), typeof(SlideItem), new PropertyMetadata(null));
	}
}
