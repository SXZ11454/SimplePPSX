using System;
using System.IO;

namespace SimplePPSX
{
	// Token: 0x02000007 RID: 7
	/// <summary>
	/// Works out which presentation app is running (PowerPoint or WPS Presentation),
	/// and from it which toolbar layout to use.
	///
	/// WPS is supported, with two documented gaps that its COM surface forces on us:
	///   - jumping to a slide by index (<see cref="PptOperator.GotoSlide"/>) does not
	///     work, so the thumbnail panels are disabled there,
	///   - clearing ink walks Shapes looking for msoInk, which WPS does not expose
	///     the same way, so the clear-all eraser action is unavailable there too.
	/// Everything else - page up/down, pen, eraser, pointer - goes through the same
	/// code for both.
	/// </summary>
	public static class PptVersionManager
	{
		// Token: 0x17000005 RID: 5
		// (get) Token: 0x06000011 RID: 17 RVA: 0x00002538 File Offset: 0x00000738
		public static string Version
		{
			get
			{
				if (!string.IsNullOrWhiteSpace(PptVersionManager._version))
				{
					return PptVersionManager._version;
				}
				return OfficeVersion.Unknown.ToString();
			}
		}

		// Token: 0x17000006 RID: 6
		// (get) Token: 0x06000012 RID: 18 RVA: 0x00002566 File Offset: 0x00000766
		public static string SoftWareType
		{
			get
			{
				return PptVersionManager._softwareType;
			}
		}

		// Token: 0x17000007 RID: 7
		// (get) Token: 0x06000013 RID: 19 RVA: 0x0000256D File Offset: 0x0000076D
		// (set) Token: 0x06000014 RID: 20 RVA: 0x00002574 File Offset: 0x00000774
		public static bool IsOffice2016OrLater { get; private set; }

		/// <summary>
		/// Is the running app WPS Presentation rather than PowerPoint?
		///
		/// The callers use it to skip the two things WPS cannot do (jump to a slide by
		/// index, clear all ink). False while nothing is running.
		/// </summary>
		public static bool IsWps { get; private set; }

		// Token: 0x06000015 RID: 21 RVA: 0x0000257C File Offset: 0x0000077C
		public static void UpdatePptApplicationInfo(string path, string versionInfo)
		{
			object locker = PptVersionManager.Locker;
			lock (locker)
			{
				PptVersionManager.IsOffice2016OrLater = false;
				PptVersionManager.IsWps = false;
				string text = null;
				try
				{
					text = Path.GetFullPath(path).ToLower();
				}
				catch (Exception)
				{
					// An odd COM path should not take the whole app down; treat it as unknown.
					text = null;
				}
				string text2;
				if (text != null && PptVersionManager.TryGetOfficeVersion(text, versionInfo, out text2))
				{
					PptVersionManager._softwareType = "office";
					PptVersionManager._version = text2;
				}
				else if (text != null && PptVersionManager.TryGetWpsVersion(text, out text2))
				{
					PptVersionManager.IsWps = true;
					PptVersionManager._softwareType = "wps";
					PptVersionManager._version = text2;
				}
				else
				{
					PptVersionManager._version = null;
					PptVersionManager._softwareType = null;
				}
			}
		}

		/// <summary>
		/// Recognises a WPS Presentation install from its executable path.
		///
		/// WPS publishes no PowerPoint-style version number to match on, so the path is
		/// the only signal. wpp.exe is the presentation app; wpsoffice.exe and wps.exe
		/// are the office shell and the writer, and must not be mistaken for it.
		/// </summary>
		private static bool TryGetWpsVersion(string fullPath, out string version)
		{
			version = null;
			if (!fullPath.Contains("wps") && !fullPath.Contains("kingsoft"))
			{
				return false;
			}
			if (fullPath.Contains("wpsoffice") || fullPath.Contains("wps.exe") || fullPath.Contains("et.exe"))
			{
				return false;
			}
			version = OfficeVersion.WPS.ToString();
			return true;
		}

		// Token: 0x06000016 RID: 22 RVA: 0x0000260C File Offset: 0x0000080C
		private static bool TryGetOfficeVersion(string path, string versionInfo, out string version)
		{
			version = null;
			if (!path.Contains("powerpoint") && !path.Contains("microsoft"))
			{
				return false;
			}
			if (!(versionInfo == "12.0"))
			{
				if (!(versionInfo == "14.0"))
				{
					if (!(versionInfo == "15.0"))
					{
						if (!(versionInfo == "16.0"))
						{
							if (versionInfo == "17.0")
							{
								version = OfficeVersion.Office2019.ToString();
								PptVersionManager.IsOffice2016OrLater = true;
							}
						}
						else
						{
							version = OfficeVersion.Office2016.ToString();
							PptVersionManager.IsOffice2016OrLater = true;
						}
					}
					else
					{
						version = OfficeVersion.Office2013.ToString();
					}
				}
				else
				{
					version = OfficeVersion.Office2010.ToString();
				}
			}
			else
			{
				version = OfficeVersion.Office2007.ToString();
			}
			return true;
		}

		// Token: 0x04000038 RID: 56
		private static string _version;

		// Token: 0x04000039 RID: 57
		private static string _softwareType;

		// Token: 0x0400003A RID: 58
		private static readonly object Locker = new object();
	}
}
