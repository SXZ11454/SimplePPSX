using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;

namespace SimplePPSX
{
	/// <summary>
	/// Restarts the application after a crash.
	/// The original located the exe through a manufacturer-namespaced registry key, which a
	/// portable install cannot rely on; it now takes the current process's own
	/// executable path directly. The only registry key this app writes is the standard
	/// per-user Run entry (TrayIconService.RunKey), used solely for start-with-Windows.
	/// </summary>
	public class Restart
	{
		public static void SimplePPSX()
		{
			string exePath = Restart.GetExecutablePath();
			if (string.IsNullOrEmpty(exePath))
			{
				return;
			}
			if (!File.Exists(exePath))
			{
				return;
			}
			Process.Start(exePath, Restart.AppRestart);
		}

		private static string GetExecutablePath()
		{
			try
			{
				Process main = Process.GetCurrentProcess();
				if (main.MainModule != null && !string.IsNullOrEmpty(main.MainModule.FileName))
				{
					return main.MainModule.FileName;
				}
			}
			catch (Exception)
			{
			}
			try
			{
				return Assembly.GetEntryAssembly().Location;
			}
			catch (Exception)
			{
				return null;
			}
		}

		private static string AppRestart = "Restart";
	}
}
