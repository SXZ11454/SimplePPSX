using System;
using System.IO;

namespace SimplePPSX
{
	// Token: 0x0200003A RID: 58
	public class ConfigHelper
	{
		// Token: 0x17000056 RID: 86
		// (get) Token: 0x0600025A RID: 602 RVA: 0x0000AEFD File Offset: 0x000090FD
		// (set) Token: 0x0600025B RID: 603 RVA: 0x0000AF04 File Offset: 0x00009104
		public static ConfigHelper Instance { get; set; }

		// Token: 0x0600025C RID: 604 RVA: 0x0000AF0C File Offset: 0x0000910C
		//
		// The constructor this used to have ran a one-time hardware check that existed
		// only to default a hardware-type flag. That integration is gone and the flag is
		// gone with it, so there is nothing left to initialise here and the constructor
		// has been removed - an implicit parameterless one serves `new ConfigHelper()`
		// in App.StartProcess.
		//
		// Nothing about "everything on by default" moved with it. IsNeedStart already
		// defaults to true in two independent places: ConfigController's constructor
		// writes it when the config file does not exist yet, and the getter passes true
		// as its fallback. So the app comes up fully enabled whether or not this
		// constructor runs.

		// Token: 0x0600025E RID: 606 RVA: 0x0000AF6A File Offset: 0x0000916A
		public bool IsNeedStart()
		{
			return ConfigHelper.ConfigController.IsNeedStart;
		}

		// Token: 0x17000057 RID: 87
		// (get) Token: 0x0600025F RID: 607 RVA: 0x0000AF76 File Offset: 0x00009176
		// (set) Token: 0x06000260 RID: 608 RVA: 0x0000AF82 File Offset: 0x00009182
		public bool IsInitApplicationSuccessful
		{
			get
			{
				return ConfigHelper.ConfigController.IsInitApplicationSuccessful;
			}
			set
			{
				ConfigHelper.ConfigController.IsInitApplicationSuccessful = value;
			}
		}

		// Token: 0x17000058 RID: 88
		// (get) Token: 0x06000261 RID: 609 RVA: 0x0000AF8F File Offset: 0x0000918F
		// (set) Token: 0x06000262 RID: 610 RVA: 0x0000AF9B File Offset: 0x0000919B
		public bool IsComOk
		{
			get
			{
				return ConfigHelper.ConfigController.IsComOk;
			}
			set
			{
				ConfigHelper.ConfigController.IsComOk = value;
			}
		}

		// Token: 0x17000059 RID: 89
		// (get) Token: 0x06000263 RID: 611 RVA: 0x0000AFA8 File Offset: 0x000091A8
		// (set) Token: 0x06000264 RID: 612 RVA: 0x0000AFB4 File Offset: 0x000091B4
		public string OfficeVersion
		{
			get
			{
				return ConfigHelper.ConfigController.OfficeVersion;
			}
			set
			{
				ConfigHelper.ConfigController.OfficeVersion = value;
			}
		}

		// Token: 0x1700005A RID: 90
		// (get) Token: 0x06000265 RID: 613 RVA: 0x0000AFC1 File Offset: 0x000091C1
		// (set) Token: 0x06000266 RID: 614 RVA: 0x0000AFCD File Offset: 0x000091CD
		public bool IsShowing
		{
			get
			{
				return ConfigHelper.ConfigController.IsShowing;
			}
			set
			{
				ConfigHelper.ConfigController.IsShowing = value;
			}
		}

		// Token: 0x1700005B RID: 91
		// (get) Token: 0x06000267 RID: 615 RVA: 0x0000AFDA File Offset: 0x000091DA
		// (set) Token: 0x06000268 RID: 616 RVA: 0x0000AFE6 File Offset: 0x000091E6
		public bool IsPptShowTypeSpeaker
		{
			get
			{
				return ConfigHelper.ConfigController.IsPptShowTypeSpeaker;
			}
			set
			{
				ConfigHelper.ConfigController.IsPptShowTypeSpeaker = value;
			}
		}

		/// <summary>
		/// Whether the overlay window is built and presented at start-up instead of
		/// at the first presentation. Static because the setting is read before any
		/// instance of this class exists - App.StartProcess builds it later on.
		/// Defaults to on; see ConfigController.IsPreload for why.
		/// </summary>
		public static bool IsPreload
		{
			get
			{
				return ConfigHelper.ConfigController.IsPreload;
			}
			set
			{
				ConfigHelper.ConfigController.IsPreload = value;
			}
		}

		/// <summary>
		/// Whether the thumbnail pass runs as soon as a presentation starts
		/// instead of on the panel's first open. Static for the same reason as
		/// <see cref="IsPreload"/>. Defaults to on; see ConfigController.IsPrefetch.
		/// </summary>
		public static bool IsPrefetch
		{
			get
			{
				return ConfigHelper.ConfigController.IsPrefetch;
			}
			set
			{
				ConfigHelper.ConfigController.IsPrefetch = value;
			}
		}

		/// <summary>
		/// The stored colour mode - system / light / dark. Static for the same
		/// reason as <see cref="IsPreload"/>. Applying it is ToolbarTheme's job.
		/// </summary>
		public static string ThemeMode
		{
			get
			{
				return ConfigHelper.ConfigController.ThemeMode;
			}
			set
			{
				ConfigHelper.ConfigController.ThemeMode = value;
			}
		}

		/// <summary>
		/// The height tier the thumbnail EXPORT pass renders its JPG files at -
		/// 270 / 540 / 1080 / 2160, or 0 for the deck's own size. Static for the
		/// same reason as <see cref="IsPreload"/>; PptAnalyzer reads it when a
		/// presentation is analysed. Screen decoding never follows this - the
		/// panel decodes at display size whatever lands on disk.
		/// </summary>
		public static int ThumbnailExportHeight
		{
			get
			{
				return ConfigHelper.ConfigController.ThumbnailExportHeight;
			}
			set
			{
				ConfigHelper.ConfigController.ThumbnailExportHeight = value;
			}
		}

		/// <summary>
		/// Whether the app writes logs at all. Static like the rest, read by
		/// <see cref="Log"/> on every write and touched only by the settings page.
		/// On by default.
		/// </summary>
		public static bool IsLogEnabled
		{
			get
			{
				return ConfigHelper.ConfigController.LogEnabled;
			}
			set
			{
				ConfigHelper.ConfigController.LogEnabled = value;
			}
		}

		/// <summary>
		/// The size a log file may reach before the next start-up clears it, in
		/// megabytes; 0 turns the auto-clean off. Checked once at start-up.
		/// </summary>
		public static int LogAutoCleanThreshold
		{
			get
			{
				return ConfigHelper.ConfigController.LogAutoCleanThreshold;
			}
			set
			{
				ConfigHelper.ConfigController.LogAutoCleanThreshold = value;
			}
		}

		// Token: 0x04000126 RID: 294
		private static readonly string Section = "StartConfig";

		// CheckedFalse is gone. It was the "false" marker written into the removed hardware
		// flag's config key, and nothing else ever used it. (Its literal was also
		// misspelled - "ChceckedFalse" - which is how it survived being the only writer
		// of its own value.)

		// Token: 0x04000129 RID: 297
		private static readonly string AppConfigFileName = "app.config";

		// Token: 0x0400012A RID: 298
		private static readonly ConfigController ConfigController = new ConfigController(Path.Combine(AppPath.RootPath(), ConfigHelper.AppConfigFileName), ConfigHelper.Section);
	}
}
