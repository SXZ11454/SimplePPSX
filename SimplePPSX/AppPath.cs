using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace SimplePPSX
{
	// Token: 0x02000032 RID: 50
	public class AppPath
	{
		// Token: 0x060001C7 RID: 455 RVA: 0x00007B68 File Offset: 0x00005D68
		public static string DocumentPath()
		{
			string text = Path.Combine(AppPath.BasePath(), AppPath.ProgramFolder, AppPath.AppName, AppPath.Document);
			if (!Directory.Exists(text))
			{
				Directory.CreateDirectory(text);
			}
			return text;
		}

		// Token: 0x060001C8 RID: 456 RVA: 0x00007BA4 File Offset: 0x00005DA4
		public static string LogPath()
		{
			string text = Path.Combine(AppPath.BasePath(), AppPath.ProgramFolder, AppPath.AppName, AppPath.Log);
			if (!Directory.Exists(text))
			{
				Directory.CreateDirectory(text);
			}
			return text;
		}

		// Token: 0x060001C9 RID: 457 RVA: 0x00007BE0 File Offset: 0x00005DE0
		public static string TemPath()
		{
			string text = Path.Combine(AppPath.BasePath(), AppPath.ProgramFolder, AppPath.AppName, AppPath.Temporary);
			if (!Directory.Exists(text))
			{
				Directory.CreateDirectory(text);
			}
			return text;
		}

		// Token: 0x060001CA RID: 458 RVA: 0x00007C1C File Offset: 0x00005E1C
		public static string RootPath()
		{
			string text = Path.Combine(AppPath.BasePath(), AppPath.ProgramFolder, AppPath.AppName);
			if (!Directory.Exists(text))
			{
				Directory.CreateDirectory(text);
			}
			return text;
		}

		// Token: 0x060001CB RID: 459 RVA: 0x00007C50 File Offset: 0x00005E50
		public static void ClearTempFolder()
		{
			try
			{
				Task.Factory.StartNew(delegate
				{
					AppPath.ClearDirectory(AppPath.TemPath());
				});
			}
			catch
			{
			}
		}

		// Token: 0x060001CC RID: 460 RVA: 0x00007C9C File Offset: 0x00005E9C
		private static void ClearDirectory(string directory)
		{
			try
			{
				Stack<string> stack = new Stack<string>();
				stack.Push(directory);
				while (stack.Any<string>())
				{
					string text = stack.Pop();
					foreach (string text2 in Directory.GetFiles(text))
					{
						File.SetAttributes(text2, FileAttributes.Normal);
						File.Delete(text2);
					}
					foreach (string text3 in Directory.GetDirectories(text))
					{
						stack.Push(text3);
					}
				}
			}
			catch
			{
			}
			try
			{
				Directory.Delete(directory, true);
				Directory.CreateDirectory(directory);
			}
			catch
			{
			}
		}

		/// <summary>
		/// The first folder of the data path, so everything lives in
		/// &lt;exe folder&gt;\Program\SimplePPSX\... : the log folder, the config file,
		/// the thumbnail cache and the temp folder.
		///
		/// Renamed from the manufacturer name it used to carry, which was both the field name and
		/// the folder. The old directory is NOT migrated: it is left where it is, and the
		/// first run after this change creates a new empty Program folder - so the app
		/// comes up with default settings and an empty log. Delete the old folder by
		/// hand once the old settings are no longer wanted.
		/// </summary>
		private static readonly string ProgramFolder = "Program";

		// Token: 0x040000C5 RID: 197
		private static readonly string AppName = "SimplePPSX";

		// Token: 0x040000C6 RID: 198
		private static readonly string Log = "Log";

		// Token: 0x040000C7 RID: 199
		private static readonly string Temporary = "Temp";

		// Token: 0x040000C8 RID: 200
		private static readonly string Document = "Document";

		/// <summary>
		/// 绿色版：所有配置与数据都放在程序自身所在目录，不再写入 %AppData%，
		/// 这样整个目录拷走即可使用，也不会在系统盘留下痕迹。
		/// </summary>
		private static string BasePath()
		{
			string baseDir = AppDomain.CurrentDomain.BaseDirectory;
			if (string.IsNullOrEmpty(baseDir))
			{
				baseDir = AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
			}
			return baseDir;
		}
	}
}
