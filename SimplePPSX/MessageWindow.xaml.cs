using System;
using System.CodeDom.Compiler;
using System.ComponentModel;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Markup;
using System.Windows.Media.Animation;

namespace SimplePPSX
{
	// Token: 0x02000035 RID: 53
	public partial class MessageWindow : Window
	{
		// Token: 0x060001F3 RID: 499 RVA: 0x00009354 File Offset: 0x00007554
		public MessageWindow(string message)
		{
			this.InitializeComponent();
			this.Message = message;
			base.DataContext = this;
			base.Loaded += this.MessageWindow_Loaded;
			base.Top = ViewHelper.MainWindow.Top;
			base.Left = ViewHelper.MainWindow.Left;
			base.Width = ViewHelper.MainWindow.Width;
			base.Height = ViewHelper.MainWindow.Height;
		}

		// Token: 0x060001F4 RID: 500 RVA: 0x000093CD File Offset: 0x000075CD
		private void MessageWindow_Loaded(object sender, RoutedEventArgs e)
		{
			this.AdjustControls();
			Storyboard storyboard = base.Resources["Style.ShowMessage"] as Storyboard;
			if (storyboard == null)
			{
				return;
			}
			storyboard.Begin();
		}

		// Token: 0x17000043 RID: 67
		// (get) Token: 0x060001F5 RID: 501 RVA: 0x000093F4 File Offset: 0x000075F4
		// (set) Token: 0x060001F6 RID: 502 RVA: 0x00009406 File Offset: 0x00007606
		public string Message
		{
			get
			{
				return (string)base.GetValue(MessageWindow.MessageProperty);
			}
			set
			{
				base.SetValue(MessageWindow.MessageProperty, value);
			}
		}

		// Token: 0x060001F7 RID: 503 RVA: 0x00009414 File Offset: 0x00007614
		private void Timeline_OnCompleted(object sender, EventArgs e)
		{
			base.Close();
		}

		// Token: 0x060001F8 RID: 504 RVA: 0x0000941C File Offset: 0x0000761C
		private void AdjustControls()
		{
			if (base.Width > 3800.0)
			{
				this.PerchGrid.Height = 140.0;
				this.MessageTextBlock.FontSize = 32.0;
			}
			if (base.Width > 1440.0 && base.Width <= 3800.0)
			{
				this.PerchGrid.Height = 80.0;
				this.MessageTextBlock.FontSize = 18.0;
			}
			if (base.Width > 0.0 && base.Width <= 1400.0)
			{
				this.PerchGrid.Height = 50.0;
				this.MessageTextBlock.FontSize = 12.0;
			}
		}

		// Token: 0x040000ED RID: 237
		public static readonly DependencyProperty MessageProperty = DependencyProperty.Register("Message", typeof(string), typeof(MessageWindow), new PropertyMetadata(null));
	}
}
