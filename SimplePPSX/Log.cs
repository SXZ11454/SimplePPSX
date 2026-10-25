using System;
using System.IO;
using System.Text;

namespace SimplePPSX
{
	// Token: 0x02000028 RID: 40
	public static class Log
	{
		/// <summary>
		/// Whether logging is on at all - the user's switch, read straight from
		/// the config each time. A write that misses at the gate costs almost
		/// nothing: the ini read is a tiny file, and the write paths here are
		/// seconds apart at the tightest (the heartbeat), never in a hot loop.
		/// </summary>
		public static bool IsEnabled
		{
			get
			{
				return ConfigHelper.IsLogEnabled;
			}
		}

		// Token: 0x0600015C RID: 348 RVA: 0x0000578F File Offset: 0x0000398F
		public static void WriteError(string error)
		{
			// The user's switch gates everything, both files. It is their choice:
			// log files are diagnostic aids and some people would rather have none
			// at all than one they never intend to read.
			if (!Log.IsEnabled)
			{
				return;
			}
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
			if (!Log.IsEnabled)
			{
				return;
			}
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

		/// <summary>
		/// Deletes both log files outright. The settings page's "clean now" is
		/// this. A locked file (a second instance, an editor holding it open) is
		/// left where it stands and the caller is told something was not removed,
		/// which is all the user needs to know.
		/// </summary>
		public static bool CleanAll()
		{
			bool removed = true;
			foreach (string path in new string[] { Log.ErrorFilePath, Log.DebugFilePath })
			{
				try
				{
					if (File.Exists(path))
					{
						File.Delete(path);
					}
				}
				catch (Exception)
				{
					// deleting a file that some process is holding is not a reason
					// to fail - the other file still gets its answer
					removed = false;
				}
			}
			return removed;
		}

		/// <summary>
		/// Startup maintenance: any log file grown past the user's retention
		/// threshold is deleted whole, so the next run starts on fresh paper.
		///
		/// The threshold is a byte cap, not an age, and that is not an assorted
		/// choice: these files append forever, never per session, so
		/// LastWriteTime is always "this run" and a day-based rule could never
		/// fire. Only size can leak in a way a user can see.
		/// </summary>
		public static void AutoCleanIfNeeded()
		{
			int thresholdMB = ConfigHelper.LogAutoCleanThreshold;
			if (thresholdMB <= 0)
			{
				return;
			}
			long thresholdBytes = (long)thresholdMB * 1024L * 1024L;
			foreach (string path in new string[] { Log.ErrorFilePath, Log.DebugFilePath })
			{
				try
				{
					if (File.Exists(path) && new FileInfo(path).Length > thresholdBytes)
					{
						File.Delete(path);
					}
				}
				catch (Exception)
				{
					// stay silent, same rule as every other log-side failure
				}
			}
		}

		/// <summary>One-line size report for the settings page, or null with no files.</summary>
		public static string DescribeSize()
		{
			long total = 0;
			foreach (string path in new string[] { Log.ErrorFilePath, Log.DebugFilePath })
			{
				try
				{
					if (File.Exists(path))
					{
						total += new FileInfo(path).Length;
					}
				}
				catch (Exception)
				{
					// unreadable now means "bigger than zero but unknown" - report 0
					// and let the next refresh tell the truth
				}
			}
			if (total <= 0L)
			{
				return string.Empty;
			}
			if (total < 1024L * 1024L)
			{
				return string.Format("当前日志占用 {0} KB", Math.Max(1, (int)(total / 1024L)));
			}
			return string.Format("当前日志占用 {0:F1} MB", (double)total / (1024.0 * 1024.0));
		}
	}
}
