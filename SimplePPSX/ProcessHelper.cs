using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace SimplePPSX
{
	// Token: 0x0200001B RID: 27
	internal static class ProcessHelper
	{
		// Looks for powerpnt.exe (PowerPoint) or wpp.exe (WPS Presentation)
		/// by walking the process list.
		///
		/// Both count as "a presentation app is running". They are never running at
		/// the same time in practice, so the first match wins.
		// Token: 0x060000E9 RID: 233 RVA: 0x00004710 File Offset: 0x00002910
		public static bool HasPowerPointProcess()
		{
			IntPtr intPtr = IntPtr.Zero;
			try
			{
				ProcessHelper.PROCESSENTRY32 processentry = new ProcessHelper.PROCESSENTRY32
				{
					dwSize = (uint)Marshal.SizeOf(typeof(ProcessHelper.PROCESSENTRY32))
				};
				intPtr = ProcessHelper.CreateToolhelp32Snapshot(2U, 0U);
				if (!ProcessHelper.Process32First(intPtr, ref processentry))
				{
					throw new ApplicationException(string.Format("Failed with win32 error code {0}", Marshal.GetLastWin32Error()));
				}
				for (;;)
				{
					string name = processentry.szExeFile;
					if (name != null)
					{
						name = name.ToLower();
						if (name == ProcessHelper.PptName || name == ProcessHelper.WppName)
						{
							return true;
						}
					}
					if (!ProcessHelper.Process32Next(intPtr, ref processentry))
					{
						return false;
					}
				}
			}
			catch (Exception ex)
			{
				throw new ApplicationException("Can't get the process.", ex);
			}
			finally
			{
				ProcessHelper.CloseHandle(intPtr);
			}
		}

		/// <summary>
		/// Is a WPS Presentation (wpp.exe) running right now?
		///
		/// Used to decide whether a stale wpp.exe may be closed: WPS leaves the process
		/// alive after its last presentation is closed, unlike PowerPoint.
		/// </summary>
		public static bool HasWpsProcess()
		{
			IntPtr intPtr = IntPtr.Zero;
			try
			{
				ProcessHelper.PROCESSENTRY32 processentry = new ProcessHelper.PROCESSENTRY32
				{
					dwSize = (uint)Marshal.SizeOf(typeof(ProcessHelper.PROCESSENTRY32))
				};
				intPtr = ProcessHelper.CreateToolhelp32Snapshot(2U, 0U);
				if (!ProcessHelper.Process32First(intPtr, ref processentry))
				{
					return false;
				}
				for (;;)
				{
					string name = processentry.szExeFile;
					if (name != null && name.ToLower() == ProcessHelper.WppName)
					{
						return true;
					}
					if (!ProcessHelper.Process32Next(intPtr, ref processentry))
					{
						return false;
					}
				}
			}
			catch (Exception)
			{
				// Not being able to look is not a reason to skip the cleanup path.
				return true;
			}
			finally
			{
				ProcessHelper.CloseHandle(intPtr);
			}
		}

		/// <summary>
		/// Closes a WPS Presentation that is running with no presentation open.
		///
		/// WPS does not quit on its own the way PowerPoint does, and the leftover
		/// process would keep this app attached to a dead COM object. Failure is not
		/// worth reporting: the caller only wants the process gone.
		/// </summary>
		public static void TryKillWppProcess()
		{
			try
			{
				foreach (string name in new string[] { "wpp" })
				{
					foreach (Process process in Process.GetProcessesByName(name))
					{
						try
						{
							if (!process.HasExited)
							{
								process.Kill();
								process.WaitForExit(2000);
							}
						}
						catch (Exception)
						{
							// it exited on its own, or we may not touch it: either is fine
						}
						finally
						{
							process.Dispose();
						}
					}
				}
			}
			catch (Exception)
			{
			}
		}

		// Token: 0x060000EB RID: 235
		[DllImport("kernel32", CharSet = CharSet.Auto, SetLastError = true)]
		private static extern IntPtr CreateToolhelp32Snapshot([In] uint dwFlags, [In] uint th32ProcessID);

		// Token: 0x060000EC RID: 236
		[DllImport("kernel32", CharSet = CharSet.Auto, SetLastError = true)]
		private static extern bool Process32First([In] IntPtr hSnapshot, ref ProcessHelper.PROCESSENTRY32 lppe);

		// Token: 0x060000ED RID: 237
		[DllImport("kernel32", CharSet = CharSet.Auto, SetLastError = true)]
		private static extern bool Process32Next([In] IntPtr hSnapshot, ref ProcessHelper.PROCESSENTRY32 lppe);

		// Token: 0x060000EE RID: 238
		[DllImport("kernel32", SetLastError = true)]
		[return: MarshalAs(UnmanagedType.Bool)]
		private static extern bool CloseHandle([In] IntPtr hObject);

		// Token: 0x04000049 RID: 73
		private const string PptName = "powerpnt.exe";

		// WPS Presentation. Same role as PptName above.
		private const string WppName = "wpp.exe";

		// Token: 0x0200008D RID: 141
		[Flags]
		private enum SnapshotFlags : uint
		{
			// Token: 0x04000230 RID: 560
			HeapList = 1U,
			// Token: 0x04000231 RID: 561
			Process = 2U,
			// Token: 0x04000232 RID: 562
			Thread = 4U,
			// Token: 0x04000233 RID: 563
			Module = 8U,
			// Token: 0x04000234 RID: 564
			Module32 = 16U,
			// Token: 0x04000235 RID: 565
			Inherit = 2147483648U,
			// Token: 0x04000236 RID: 566
			All = 31U,
			// Token: 0x04000237 RID: 567
			NoHeaps = 1073741824U
		}

		// Token: 0x0200008E RID: 142
		[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
		private struct PROCESSENTRY32
		{
			// Token: 0x04000238 RID: 568
			private const int MAX_PATH = 260;

			// Token: 0x04000239 RID: 569
			internal uint dwSize;

			// Token: 0x0400023A RID: 570
			internal uint cntUsage;

			// Token: 0x0400023B RID: 571
			internal uint th32ProcessID;

			// Token: 0x0400023C RID: 572
			internal IntPtr th32DefaultHeapID;

			// Token: 0x0400023D RID: 573
			internal uint th32ModuleID;

			// Token: 0x0400023E RID: 574
			internal uint cntThreads;

			// Token: 0x0400023F RID: 575
			internal uint th32ParentProcessID;

			// Token: 0x04000240 RID: 576
			internal int pcPriClassBase;

			// Token: 0x04000241 RID: 577
			internal uint dwFlags;

			// Token: 0x04000242 RID: 578
			[MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
			internal string szExeFile;
		}
	}
}
