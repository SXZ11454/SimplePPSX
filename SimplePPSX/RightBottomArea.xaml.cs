using System;
using System.Windows;
using System.Windows.Controls;

namespace SimplePPSX
{
	/// <summary>
	/// Right toolbar, horizontal variant (Office 2016 and newer). Show and hide are
	/// MainWindow's job (ToolbarAnimation).
	/// </summary>
	public partial class RightBottomArea : UserControl
	{
		public RightBottomArea()
		{
			this.InitializeComponent();
			// Backdrop and hairlines join the toolbar chrome that the opacity
			// setting drives (see ToolbarAppearance).
			ToolbarAppearance.Register(this.ChromeRectangle);
			// Hovering any part of this toolbar raises the backdrop to the default
			// opacity; leaving glides it back to the idle ("not clicked") one.
			ToolbarAppearance.HookHover(this);
		}

		private void Previous_Click(object sender, RoutedEventArgs e)
		{
			PptOperator.Instance.PreviousPage();
			MainWindow mainWindow = ViewHelper.MainWindow;
			if (mainWindow != null)
			{
				mainWindow.HideAllPanels();
			}
			PptOperator.Instance.BringPresentationToFront();
		}

		private void Next_Click(object sender, RoutedEventArgs e)
		{
			PptOperator.Instance.NextPage();
			MainWindow mainWindow = ViewHelper.MainWindow;
			if (mainWindow != null)
			{
				mainWindow.HideAllPanels();
			}
			PptOperator.Instance.BringPresentationToFront();
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
