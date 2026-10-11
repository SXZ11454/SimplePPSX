using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace SimplePPSX
{
	/// <summary>
	/// Width of one of the two columns in a thumbnail entry template
	/// (Style_\ToolbarStyle.xaml, Style.ThumbnailItem*).
	///
	/// An entry is [number column][image column]: the number side is a fixed
	/// column so the images line up down the list, and the other side is a
	/// star that resolves to whatever the entry has left - which is the
	/// image's own width, since the entry is number column plus image. The
	/// right panel's template moves the IMAGE into the other column (its
	/// DataTrigger sets Grid.Column on the image host and on the number),
	/// and the two column widths have to move with it or the picture lands
	/// in the fixed 36px column and is squeezed to a sliver.
	///
	/// The widths cannot stay plain bindings with a trigger setter beside
	/// them: a data binding outranks a template trigger in WPF's value
	/// precedence, so a trigger can never override a bound width. Both
	/// columns therefore bind through this converter, which reads the
	/// number-column width and the panel's IsLeftSide and picks the side -
	/// "flip" (ConverterParameter) produces the mirror image of the answer,
	/// for the column the number does not stand on.
	/// </summary>
	public sealed class ThumbnailColumnWidthConverter : IMultiValueConverter
	{
		/// <summary>
		/// Stands in when the number-column dependency property has not been
		/// measured yet: the middle tier's number column, the same value
		/// MainWindow's property metadata carries as its default.
		/// </summary>
		private const double FallbackNumberColumn = 36.0;

		public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
		{
			double numberColumn = ThumbnailColumnWidthConverter.FallbackNumberColumn;
			if (values != null && values.Length > 0 && values[0] is double)
			{
				double measured = (double)values[0];
				if (!double.IsNaN(measured) && measured > 0.0)
				{
					numberColumn = measured;
				}
			}

			// True - the number stands left of the image - is the default, and
			// also the answer when IsLeftSide has not resolved yet: the left
			// panel is the one whose entries must look right on the first
			// frame, before either binding has travelled.
			bool isLeftSide = true;
			if (values != null && values.Length > 1 && values[1] is bool)
			{
				isLeftSide = (bool)values[1];
			}

			if (string.Equals(parameter as string, "flip", StringComparison.Ordinal))
			{
				isLeftSide = !isLeftSide;
			}

			return isLeftSide
				? new GridLength(numberColumn)
				: new GridLength(1.0, GridUnitType.Star);
		}

		public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
		{
			throw new NotSupportedException("Thumbnail column widths are one-way.");
		}
	}
}
