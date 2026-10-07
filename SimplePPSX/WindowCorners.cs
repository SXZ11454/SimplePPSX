using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;

namespace SimplePPSX
{
	/// <summary>
	/// Windows 11 rounds the corners of every top level window by itself - the last
	/// place that does not match the square-corner design of the rest of the app
	/// (Style_\SquareCorners.xaml). Ask the desktop window manager not to round them.
	///
	/// Call <see cref="DisableRounding"/> once per window, after its handle exists
	/// (SourceInitialized). Older Windows versions do not know the attribute: the call
	/// fails, the result is ignored and nothing changes.
	/// </summary>
	internal static class WindowCorners
	{
		/// <summary>DWMWA_WINDOW_CORNER_PREFERENCE (Windows 11 and newer)</summary>
		private const int DwmwaWindowCornerPreference = 33;

		/// <summary>DWMWCP_DONOTROUND</summary>
		private const int CornerPreferenceDontRound = 1;

		[DllImport("dwmapi.dll", PreserveSig = true)]
		private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attribute, ref int value, int valueSize);

		/// <summary>Turns off the automatic rounding of one window's corners.</summary>
		internal static void DisableRounding(Window window)
		{
			IntPtr handle = new WindowInteropHelper(window).Handle;
			if (handle == IntPtr.Zero)
			{
				return; // no handle yet: nothing to talk to
			}

			int preference = CornerPreferenceDontRound;
			DwmSetWindowAttribute(handle, DwmwaWindowCornerPreference, ref preference, sizeof(int));
		}
	}
}
