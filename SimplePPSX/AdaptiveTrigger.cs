using System;
using System.ComponentModel;
using System.Windows;
using Microsoft.Xaml.Behaviors;

namespace SimplePPSX
{
	/// <summary>
	/// Fires the visual states of the element it is attached to once the owning
	/// window reaches a given width or height, which is how the toolbars switch
	/// between their three size tiers.
	///
	/// This used to derive from the Blend SDK's System.Windows.Interactivity
	/// TriggerBase, which is a .NET Framework only package. The supported
	/// replacement (Microsoft.Xaml.Behaviors.Wpf) exposes the same base class and
	/// the same OnAttached / OnDetaching / InvokeActions members, so only the using
	/// changed.
	/// </summary>
	public class AdaptiveTrigger : TriggerBase<FrameworkElement>
	{
		/// <summary>
		/// The window is not on screen yet, so the size is taken on Loaded.
		/// </summary>
		protected override void OnAttached()
		{
			if (DesignerProperties.GetIsInDesignMode(base.AssociatedObject))
			{
				return;
			}
			Window window = Window.GetWindow(base.AssociatedObject);
			if (window == null)
			{
				base.AssociatedObject.Loaded += this.OnLoaded;
				return;
			}
			if (window.ActualWidth <= 0.0 && window.ActualHeight <= 0.0)
			{
				base.AssociatedObject.Loaded += this.OnLoaded;
				return;
			}
			this.StartWatch();
		}

		protected override void OnDetaching()
		{
			if (DesignerProperties.GetIsInDesignMode(base.AssociatedObject))
			{
				return;
			}
			base.AssociatedObject.Loaded -= this.OnLoaded;
			Window window = Window.GetWindow(base.AssociatedObject);
			if (window != null)
			{
				window.SizeChanged -= this.OnWindowSizeChanged;
			}
		}

		private void OnLoaded(object sender, RoutedEventArgs routedEventArgs)
		{
			base.AssociatedObject.Loaded -= this.OnLoaded;
			this.StartWatch();
		}

		private void OnWindowSizeChanged(object sender, SizeChangedEventArgs e)
		{
			this.Adapt(e.NewSize);
		}

		private void StartWatch()
		{
			Window window = Window.GetWindow(base.AssociatedObject);
			if (window == null)
			{
				return;
			}
			window.SizeChanged += this.OnWindowSizeChanged;
			if (base.AssociatedObject.IsLoaded)
			{
				this.Adapt(new Size(window.ActualWidth, window.ActualHeight));
				return;
			}
			base.Dispatcher.BeginInvoke(new Action(delegate
			{
				this.Adapt(new Size(window.ActualWidth, window.ActualHeight));
			}));
		}

		private void Adapt(Size size)
		{
			bool matched = false;
			if (!double.IsNaN(this.MinWindowWidth))
			{
				matched = size.Width >= this.MinWindowWidth;
			}
			if (!matched && !double.IsNaN(this.MinWindowHeight))
			{
				matched = size.Height >= this.MinWindowHeight;
			}
			if (matched)
			{
				base.InvokeActions(null);
			}
		}

		public double MinWindowWidth
		{
			get
			{
				return (double)base.GetValue(AdaptiveTrigger.MinWindowWidthProperty);
			}
			set
			{
				base.SetValue(AdaptiveTrigger.MinWindowWidthProperty, value);
			}
		}

		public double MinWindowHeight
		{
			get
			{
				return (double)base.GetValue(AdaptiveTrigger.MinWindowHeightProperty);
			}
			set
			{
				base.SetValue(AdaptiveTrigger.MinWindowHeightProperty, value);
			}
		}

		public static readonly DependencyProperty MinWindowWidthProperty = DependencyProperty.Register("MinWindowWidth", typeof(double), typeof(AdaptiveTrigger), new PropertyMetadata(double.NaN));

		public static readonly DependencyProperty MinWindowHeightProperty = DependencyProperty.Register("MinWindowHeight", typeof(double), typeof(AdaptiveTrigger), new PropertyMetadata(double.NaN));
	}
}
