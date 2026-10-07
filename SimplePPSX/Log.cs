using System;
using System.IO;
using System.Text;

namespace SimplePPSX
{
	// Token: 0x02000028 RID: 40
	public static class Log
	{
		// Token: 0x0600015C RID: 348 RVA: 0x0000578F File Offset: 0x0000398F
		public static void WriteError(string error)
		{
			Log.Write(Log.ErrorFilePath, error);
		}

		/// <summary>
		/// Writes a diagnostic line to debug.txt.
		///
		/// This used to have an empty body - a stub in the original code - which meant
		/// nothing the app recorded about its own state could ever be looked at. That
		/// is exactly what you need when a toolbar refuses to appear and the only
		/// question is whether the presentation app was ever heard from: there was
		/// nowhere to look.
		///
		/// Failures are swallowed. Debug output must never be able to take the app
		/// down, and a locked or full log is not worth reporting to a user who is
		/// busy presenting.
		/// </summary>
		// Token: 0x0600015D RID: 349 RVA: 0x00002F72 File Offset: 0x00001172
		public static void WriteDebug(string info)
		{
			try
			{
				Log.Write(Log.DebugFilePath, info);
			}
			catch (Exception)
			{
				// logging is not worth failing over
			}
		}

		// Token: 0x0600015E RID: 350 RVA: 0x0000579C File Offset: 0x0000399C
		private static void Write(string filePath, string str)
		{
			object locker = Log.Locker;
			lock (locker)
			{
				try
				{
					StringBuilder stringBuilder = new StringBuilder();
					stringBuilder.AppendLine(string.Format("{0:G}  {1}", DateTime.Now, str));
					File.AppendAllText(filePath, stringBuilder.ToString());
				}
				catch (Exception)
				{
					// The log is a diagnostic aid, never a reason to fail an action. The
					// error path writes through here too, so throwing from it would turn
					// a full disk or a locked file into an unhandled exception.
				}
			}
		}
		// Token: 0x0400006B RID: 107
		public static readonly string ErrorFilePath = Path.Combine(AppPath.LogPath(), "error.txt");

		// Token: 0x0400006C RID: 108
		public static readonly string DebugFilePath = Path.Combine(AppPath.LogPath(), "debug.txt");

		// Token: 0x0400006D RID: 109
		private static readonly object Locker = new object();
	}
}
