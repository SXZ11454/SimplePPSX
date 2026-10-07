using System;
using System.Diagnostics;
using System.Windows;
using Microsoft.Office.Interop.PowerPoint;

namespace SimplePPSX
{
	// Token: 0x02000016 RID: 22
	public class Gesture
	{
		// Token: 0x060000C6 RID: 198 RVA: 0x00004004 File Offset: 0x00002204
		public Gesture()
		{
			this._stopwatch = new Stopwatch();
			ViewHelper.MainWindow.LocationChanged += delegate(object s, EventArgs e)
			{
				this.ChangeGestureScope();
			};
			ViewHelper.MainWindow.SizeChanged += delegate(object s, SizeChangedEventArgs e)
			{
				this.ChangeGestureScope();
			};
			MouseHelper.Instance.Down += this.Instance_Down;
			MouseHelper.Instance.Move += this.Instance_Move;
			MouseHelper.Instance.Up += this.Instance_Up;
		}

		// Token: 0x060000C7 RID: 199 RVA: 0x00004090 File Offset: 0x00002290
		public void ChangeGestureScope()
		{
			this._gestureScope = new Rect(ViewHelper.MainWindow.Left, ViewHelper.MainWindow.Top, ViewHelper.MainWindow.Width, ViewHelper.MainWindow.Height);
		}

		// Token: 0x060000C8 RID: 200 RVA: 0x000040C8 File Offset: 0x000022C8
		private void Instance_Down(object sender, MouseHelperEventArgs e)
		{
			if (!this._gestureScope.Contains(e.MousePoint))
			{
				return;
			}
			this._maxLengthX = 0.0;
			this._maxLengthY = 0.0;
			this._startPoint = e.MousePoint;
			this._stopwatch.Restart();
		}

		// Token: 0x060000C9 RID: 201 RVA: 0x00004120 File Offset: 0x00002320
		private void Instance_Move(object sender, MouseHelperEventArgs e)
		{
			if (!this._gestureScope.Contains(e.MousePoint))
			{
				return;
			}
			double num = Math.Abs(e.MousePoint.X - this._startPoint.X);
			double num2 = Math.Abs(e.MousePoint.Y - this._startPoint.Y);
			if (num > this._maxLengthX)
			{
				this._maxLengthX = num;
			}
			if (num2 > this._maxLengthY)
			{
				this._maxLengthY = num2;
			}
		}

		// Token: 0x060000CA RID: 202 RVA: 0x000041A4 File Offset: 0x000023A4
		private void Instance_Up(object sender, MouseHelperEventArgs e)
		{
			if (!this._gestureScope.Contains(e.MousePoint))
			{
				return;
			}
			this._stopwatch.Stop();
			if ((double)this._stopwatch.ElapsedMilliseconds > Gesture.MaxGestureTime || (double)this._stopwatch.ElapsedMilliseconds < Gesture.MinGestureTime)
			{
				return;
			}
			if (!this.HorizontalGesture(e.MousePoint))
			{
				this.VerticalGesture(e.MousePoint);
			}
		}

		// Token: 0x060000CB RID: 203 RVA: 0x00004214 File Offset: 0x00002414
		private bool HorizontalGesture(Point position)
		{
			if (this._maxLengthY > Gesture.MaxGestureLengthXorY)
			{
				return false;
			}
			double num = this._startPoint.X - position.X;
			if (num > Gesture.MinGestureLengthXorY && this.isPptArrow())
			{
				this.NextPage();
				return true;
			}
			if (num < -Gesture.MinGestureLengthXorY && this.isPptArrow())
			{
				this.PreviousPage();
				return true;
			}
			return false;
		}

		// Token: 0x060000CC RID: 204 RVA: 0x00004278 File Offset: 0x00002478
		private void VerticalGesture(Point position)
		{
			if (this._maxLengthX > Gesture.MaxGestureLengthXorY)
			{
				return;
			}
			double num = this._startPoint.Y - position.Y;
			if (num > Gesture.MinGestureLengthXorY && this.isPptArrow())
			{
				this.NextPage();
			}
			if (num < -Gesture.MinGestureLengthXorY && this.isPptArrow())
			{
				this.PreviousPage();
			}
		}

		// Token: 0x060000CD RID: 205 RVA: 0x000042D2 File Offset: 0x000024D2
		private bool isPptArrow()
		{
			return PptOperator.Instance.PointerType == PpSlideShowPointerType.ppSlideShowPointerNone || PptOperator.Instance.PointerType == PpSlideShowPointerType.ppSlideShowPointerAutoArrow || PptOperator.Instance.PointerType == PpSlideShowPointerType.ppSlideShowPointerArrow;
		}

		// Token: 0x060000CE RID: 206 RVA: 0x000042FC File Offset: 0x000024FC
		private void PreviousPage()
		{
			SyntacticSugar.TryCatchAction(delegate
			{
				if (ViewHelper.BottomArea.IsNeedStopMouseLineTrack || ViewHelper.MainWindow.IsLeftSlideManagerViewShowed || ViewHelper.MainWindow.IsRightSlideManagerViewShowed || Math.Abs(ViewHelper.MainWindow.Opacity) < 0.1)
				{
					return;
				}
				PptOperator.Instance.PreviousPage();
			}, null);
		}

		// Token: 0x060000CF RID: 207 RVA: 0x00004323 File Offset: 0x00002523
		private void NextPage()
		{
			SyntacticSugar.TryCatchAction(delegate
			{
				if (ViewHelper.BottomArea.IsNeedStopMouseLineTrack || ViewHelper.MainWindow.IsLeftSlideManagerViewShowed || ViewHelper.MainWindow.IsRightSlideManagerViewShowed || Math.Abs(ViewHelper.MainWindow.Opacity) < 0.1)
				{
					return;
				}
				PptOperator.Instance.NextPage();
			}, null);
		}

		// Token: 0x04000036 RID: 54
		private Point _startPoint;

		// Token: 0x04000037 RID: 55
		private Rect _gestureScope;

		// Token: 0x04000038 RID: 56
		private double _maxLengthX;

		// Token: 0x04000039 RID: 57
		private double _maxLengthY;

		// Token: 0x0400003A RID: 58
		private readonly Stopwatch _stopwatch;

		// Token: 0x0400003B RID: 59
		private static readonly double MinGestureTime = 100.0;

		// Token: 0x0400003C RID: 60
		private static readonly double MaxGestureTime = 600.0;

		// Token: 0x0400003D RID: 61
		private static readonly double MinGestureLengthXorY = 30.0;

		// Token: 0x0400003E RID: 62
		private static readonly double MaxGestureLengthXorY = 70.0;
	}
}
