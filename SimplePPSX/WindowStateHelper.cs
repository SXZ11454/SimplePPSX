using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Interop;

namespace SimplePPSX
{
	// Token: 0x02000020 RID: 32
	public static class WindowStateHelper
	{
		// Token: 0x0600010A RID: 266 RVA: 0x00004A94 File Offset: 0x00002C94
		public static void UpdateWindowRect(Window window)
		{
			WindowRect rect;
			WindowsApi.GetWindowRect(new HandleRef(window, PptOperator.Instance.SlideShowWindowIntPtr), out rect);
			if (Math.Abs(rect.Left - WindowStateHelper.PreWindowRect.Left) < 1 && Math.Abs(rect.Top - WindowStateHelper.PreWindowRect.Top) < 1 && Math.Abs(rect.Right - WindowStateHelper.PreWindowRect.Right) < 1 && Math.Abs(rect.Bottom - WindowStateHelper.PreWindowRect.Bottom) < 1)
			{
				return;
			}
			// Post-or-inline, never a blocking hop: the UI timer calls this on the
			// UI thread (where the hop is inline anyway), but PowerPoint's
			// show-begin callback calls it from PowerPoint's own event thread - and
			// an event thread that waits for the UI while the UI waits on a COM
			// call back into PowerPoint parks both. See
			// SyntacticSugar.DispatcherPostOrInline.
			SyntacticSugar.DispatcherPostOrInline(delegate
			{
				window.Top = (double)rect.Top;
				window.Left = (double)rect.Left / ScreenHelper.DpiFactor;
				window.Width = (double)(rect.Right - rect.Left) / ScreenHelper.DpiFactor;
				window.Height = (double)(rect.Bottom - rect.Top) / ScreenHelper.DpiFactor;
			});
			WindowStateHelper.PreWindowRect = rect;
		}

		// Token: 0x0600010B RID: 267 RVA: 0x00004B64 File Offset: 0x00002D64
		public static bool GetPptShowingState(Window window)
		{
			WindowStateHelper.Sb.Clear();
			List<IntPtr> childWindows = WindowsApi.GetChildWindows(PptOperator.Instance.SlideShowWindowIntPtr);
			childWindows.Add(new WindowInteropHelper(window).Handle);
			childWindows.Add(WindowStateHelper.LoadingWindowIntPtr);
			IntPtr intPtr = WindowsApi.WindowFromPoint(new MousePoint(window.Left, window.Top));
			IntPtr intPtr2 = WindowsApi.WindowFromPoint(new MousePoint(window.Left + window.Width - (double)WindowStateHelper.Offset, window.Top + (double)WindowStateHelper.Offset));
			IntPtr intPtr3 = WindowsApi.WindowFromPoint(new MousePoint(window.Left + window.Width - (double)WindowStateHelper.Offset, window.Height - (double)WindowStateHelper.Offset));
			IntPtr intPtr4 = WindowsApi.WindowFromPoint(new MousePoint(window.Left + window.Width / 2.0 - (double)WindowStateHelper.Offset, window.Top + window.Height / 2.0 - (double)WindowStateHelper.Offset));
			if (intPtr == intPtr2)
			{
				WindowsApi.GetWindowText(intPtr, WindowStateHelper.Sb, WindowStateHelper.Sb.Capacity);
				if (WindowStateHelper.Sb.ToString().ToLower() == WindowStateHelper.RemoteShow)
				{
					WindowsApi.SetWindowPos(intPtr, 0, 0, 1, 0, 0, 69);
					return true;
				}
			}
			return childWindows.Contains(intPtr4) || childWindows.Contains(intPtr) || childWindows.Contains(intPtr2) || childWindows.Contains(intPtr3);
		}

		// Token: 0x0600010C RID: 268 RVA: 0x00004CD0 File Offset: 0x00002ED0
		public static void MakeAppAbovePpt(Window window)
		{
			IntPtr intPtr = new WindowInteropHelper(window).Handle;
			for (int i = 0; i < WindowStateHelper.MaxFindWindows; i++)
			{
				intPtr = WindowsApi.GetWindow(intPtr, GetWindowCmd.GwHwndnext);
				if (intPtr == PptOperator.Instance.SlideShowWindowIntPtr)
				{
					return;
				}
			}
			WindowsApi.SetWindowPos(new WindowInteropHelper(window).Handle, -1, 0, 0, 0, 0, 3);
		}

		// Token: 0x04000056 RID: 86
		public static WindowRect PreWindowRect;

		// Token: 0x04000057 RID: 87
		private static readonly int Offset = 5;

		// Token: 0x04000058 RID: 88
		public static IntPtr LoadingWindowIntPtr;

		// Token: 0x04000059 RID: 89
		private static readonly int MaxFindWindows = 1000;

		// Token: 0x0400005A RID: 90
		private static readonly string RemoteShow = "remoteshow";

		// Token: 0x0400005B RID: 91
		private static readonly StringBuilder Sb = new StringBuilder(1024);
	}
}
