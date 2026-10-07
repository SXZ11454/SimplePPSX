using System;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;

namespace SimplePPSX
{
	// Token: 0x0200000A RID: 10
	public class ControlHelper : DependencyObject
	{
		// Token: 0x0600003F RID: 63 RVA: 0x00002981 File Offset: 0x00000B81
		public static Thickness GetContentMargin(DependencyObject obj)
		{
			return (Thickness)obj.GetValue(ControlHelper.ContentMarginProperty);
		}

		// Token: 0x06000040 RID: 64 RVA: 0x00002993 File Offset: 0x00000B93
		public static void SetContentMargin(DependencyObject obj, Thickness value)
		{
			obj.SetValue(ControlHelper.ContentMarginProperty, value);
		}

		// Token: 0x06000041 RID: 65 RVA: 0x000029A6 File Offset: 0x00000BA6
		public static CornerRadius GetCornerRadius(DependencyObject obj)
		{
			return (CornerRadius)obj.GetValue(ControlHelper.CornerRadiusProperty);
		}

		// Token: 0x06000042 RID: 66 RVA: 0x000029B8 File Offset: 0x00000BB8
		public static void SetCornerRadius(DependencyObject obj, CornerRadius value)
		{
			obj.SetValue(ControlHelper.CornerRadiusProperty, value);
		}

		// Token: 0x06000043 RID: 67 RVA: 0x000029CB File Offset: 0x00000BCB
		public static ImageSource GetNormalImage(DependencyObject obj)
		{
			return (ImageSource)obj.GetValue(ControlHelper.PressedImageProperty);
		}

		// Token: 0x06000044 RID: 68 RVA: 0x000029DD File Offset: 0x00000BDD
		public static void SetNormalImage(DependencyObject obj, ImageSource value)
		{
			obj.SetValue(ControlHelper.NormalImageProperty, value);
		}

		// Token: 0x06000045 RID: 69 RVA: 0x000029CB File Offset: 0x00000BCB
		public static ImageSource GetPressedImage(DependencyObject obj)
		{
			return (ImageSource)obj.GetValue(ControlHelper.PressedImageProperty);
		}

		// Token: 0x06000046 RID: 70 RVA: 0x000029EB File Offset: 0x00000BEB
		public static void SetPressedImage(DependencyObject obj, ImageSource value)
		{
			obj.SetValue(ControlHelper.PressedImageProperty, value);
		}

		// Token: 0x06000047 RID: 71 RVA: 0x000029F9 File Offset: 0x00000BF9
		public static ImageSource GetCheckedImage(DependencyObject obj)
		{
			return (ImageSource)obj.GetValue(ControlHelper.CheckedImageProperty);
		}

		// Token: 0x06000048 RID: 72 RVA: 0x00002A0B File Offset: 0x00000C0B
		public static void SetCheckedImage(DependencyObject obj, ImageSource value)
		{
			obj.SetValue(ControlHelper.CheckedImageProperty, value);
		}

		// Token: 0x06000049 RID: 73 RVA: 0x00002A1C File Offset: 0x00000C1C
		public static void SetAltF4Enabled(Window window, bool value)
		{
			window.SetValue(ControlHelper.AltF4EnabledProperty, value);
			if (!value)
			{
				window.PreviewKeyDown += ControlHelper.Window_PreviewKeyDown;
				window.PreviewKeyUp += ControlHelper.Window_PreviewKeyUp;
				return;
			}
			window.PreviewKeyDown -= ControlHelper.Window_PreviewKeyDown;
			window.PreviewKeyUp -= ControlHelper.Window_PreviewKeyUp;
		}

		// Token: 0x0600004A RID: 74 RVA: 0x00002A86 File Offset: 0x00000C86
		public static bool GetAltF4Enabled(Window window)
		{
			return (bool)window.GetValue(ControlHelper.AltF4EnabledProperty);
		}

		// Token: 0x0600004B RID: 75 RVA: 0x00002A98 File Offset: 0x00000C98
		private static void Window_PreviewKeyDown(object sender, KeyEventArgs e)
		{
			Window window = sender as Window;
			if (window == null)
			{
				return;
			}
			if (e.SystemKey == Key.LeftAlt || e.SystemKey == Key.RightAlt)
			{
				ControlHelper.SetIsAltDown(window, true);
				return;
			}
			if (e.SystemKey == Key.F4 && ControlHelper.GetIsAltDown(window))
			{
				e.Handled = true;
			}
		}

		// Token: 0x0600004C RID: 76 RVA: 0x00002AE8 File Offset: 0x00000CE8
		private static void Window_PreviewKeyUp(object sender, KeyEventArgs e)
		{
			Window window = sender as Window;
			if (window == null)
			{
				return;
			}
			if (e.SystemKey == Key.LeftAlt || e.SystemKey == Key.RightAlt)
			{
				ControlHelper.SetIsAltDown(window, false);
			}
		}

		// Token: 0x0600004D RID: 77 RVA: 0x00002B1B File Offset: 0x00000D1B
		private static void SetIsAltDown(DependencyObject element, bool value)
		{
			element.SetValue(ControlHelper.IsAltDownProperty, value);
		}

		// Token: 0x0600004E RID: 78 RVA: 0x00002B2E File Offset: 0x00000D2E
		private static bool GetIsAltDown(DependencyObject element)
		{
			return (bool)element.GetValue(ControlHelper.IsAltDownProperty);
		}

		// Token: 0x04000007 RID: 7
		public static readonly DependencyProperty ContentMarginProperty = DependencyProperty.RegisterAttached("ContentMargin", typeof(Thickness), typeof(ControlHelper), new PropertyMetadata(null));

		// Token: 0x04000008 RID: 8
		public static readonly DependencyProperty CornerRadiusProperty = DependencyProperty.RegisterAttached("CornerRadius", typeof(CornerRadius), typeof(ControlHelper), new PropertyMetadata(default(CornerRadius)));

		// Token: 0x04000009 RID: 9
		public static readonly DependencyProperty NormalImageProperty = DependencyProperty.RegisterAttached("NormalImage", typeof(ImageSource), typeof(ControlHelper), new PropertyMetadata(null));

		// Token: 0x0400000A RID: 10
		public static readonly DependencyProperty PressedImageProperty = DependencyProperty.RegisterAttached("PressedImage", typeof(ImageSource), typeof(ControlHelper), new PropertyMetadata(null));

		// Token: 0x0400000B RID: 11
		public static readonly DependencyProperty CheckedImageProperty = DependencyProperty.RegisterAttached("CheckedImage", typeof(ImageSource), typeof(ControlHelper), new PropertyMetadata(null));

		// Token: 0x0400000C RID: 12
		public static readonly DependencyProperty AltF4EnabledProperty = DependencyProperty.RegisterAttached("AltF4Enabled", typeof(bool), typeof(ControlHelper), new PropertyMetadata(true));

		// Token: 0x0400000D RID: 13
		private static readonly DependencyProperty IsAltDownProperty = DependencyProperty.RegisterAttached("IsAltDown", typeof(bool), typeof(ControlHelper), new PropertyMetadata(false));
	}
}
