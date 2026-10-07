using System;
using System.Windows;
using System.Windows.Controls;

namespace SimplePPSX
{
	/// <summary>
	/// Right toolbar, vertical variant: same structure as the left one. Show and hide
	/// are MainWindow's job (ToolbarAnimation).
	/// </summary>
	public partial class RightArea : UserControl
	{
		public RightArea()
		{
			this.InitializeComponent();
			// Backdrop and hairlines join the toolbar chrome that the opacity
			// setting drives (see ToolbarAppearance).
			ToolbarAppearance.Register(this.ChromeRectangle, this.DividerTop, this.DividerBottom);
		}

		private void Previous_Click(object sender, RoutedEventArgs e)
		{
			PptOperator.Instance.PreviousPage();
			MainWindow mainWindow = ViewHelper.MainWindow;
			if (mainWindow != null)
			{
				mainWindow.HideAllPanels();
			}
			WindowsApi.SetForegroundWindow(PptOperator.Instance.SlideShowWindowIntPtr);
		}

		private void Next_Click(object sender, RoutedEventArgs e)
		{
			PptOperator.Instance.NextPage();
			MainWindow mainWindow = ViewHelper.MainWindow;
			if (mainWindow != null)
			{
				mainWindow.HideAllPanels();
			}
			WindowsApi.SetForegroundWindow(PptOperator.Instance.SlideShowWindowIntPtr);
		}

		private void Thumbnail_Click(object sender, RoutedEventArgs e)
		{
			// No WPS special case here any more: the jump itself goes through
			// ComLate.CallGotoSlide, and the panel is built from late-bound COM too, so
			// the same code serves both products. The shared entry point is
			// MainWindow.ToggleThumbnailPanelFromToolbar, which used to be four
			// identical copies of a guard that made this button do nothing at all.
			MainWindow mainWindow = ViewHelper.MainWindow;
			if (mainWindow == null)
			{
				return;
			}
			mainWindow.ToggleThumbnailPanelFromToolbar(mainWindow.RightThumbnailPanel);
		}
	}
}
