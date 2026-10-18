using System;
using System.IO;

namespace SimplePPSX
{
	// Token: 0x02000027 RID: 39
	internal class ConfigController : IniSetting
	{
		// Token: 0x0600014C RID: 332 RVA: 0x0000564F File Offset: 0x0000384F
		public ConfigController(string settingPath, string section)
			: base(settingPath, section)
		{
			if (!File.Exists(settingPath))
			{
				base.SetSettingValue(ConfigController.IsNeedStartKey, true, section);
			}
		}

		// Token: 0x17000035 RID: 53
		// (get) Token: 0x0600014D RID: 333 RVA: 0x0000566E File Offset: 0x0000386E
		// (set) Token: 0x0600014E RID: 334 RVA: 0x0000567C File Offset: 0x0000387C
		public bool IsComOk
		{
			get
			{
				return base.GetSettingValue(ConfigController.IsComOkKey, true);
			}
			set
			{
				base.SetSettingValue(ConfigController.IsComOkKey, value);
			}
		}

		// Token: 0x17000036 RID: 54
		// (get) Token: 0x0600014F RID: 335 RVA: 0x0000568A File Offset: 0x0000388A
		// (set) Token: 0x06000150 RID: 336 RVA: 0x00005698 File Offset: 0x00003898
		public bool IsInitApplicationSuccessful
		{
			get
			{
				return base.GetSettingValue(ConfigController.IsInitApplicationSuccessfulKey, true);
			}
			set
			{
				base.SetSettingValue(ConfigController.IsInitApplicationSuccessfulKey, value);
			}
		}

		// Token: 0x17000037 RID: 55
		// (get) Token: 0x06000151 RID: 337 RVA: 0x000056A6 File Offset: 0x000038A6
		// (set) Token: 0x06000152 RID: 338 RVA: 0x000056B4 File Offset: 0x000038B4
		public bool IsNeedStart
		{
			get
			{
				return base.GetSettingValue(ConfigController.IsNeedStartKey, true);
			}
			set
			{
				base.SetSettingValue(ConfigController.IsNeedStartKey, value);
			}
		}

		// Token: 0x17000038 RID: 56
		// (get) Token: 0x06000153 RID: 339 RVA: 0x000056C2 File Offset: 0x000038C2
		// (set) Token: 0x06000154 RID: 340 RVA: 0x000056D4 File Offset: 0x000038D4
		public string OfficeVersion
		{
			get
			{
				return base.GetSettingValue(ConfigController.OfficeVersionKey, "");
			}
			set
			{
				base.SetSettingValue(ConfigController.OfficeVersionKey, value);
			}
		}

		// Token: 0x17000039 RID: 57
		// (get) Token: 0x06000155 RID: 341 RVA: 0x000056E2 File Offset: 0x000038E2
		// (set) Token: 0x06000156 RID: 342 RVA: 0x000056F0 File Offset: 0x000038F0
		public bool IsShowing
		{
			get
			{
				return base.GetSettingValue(ConfigController.IsShowingKey, false);
			}
			set
			{
				base.SetSettingValue(ConfigController.IsShowingKey, value);
			}
		}


		// Token: 0x1700003B RID: 59
		// (get) Token: 0x06000159 RID: 345 RVA: 0x0000571E File Offset: 0x0000391E
		// (set) Token: 0x0600015A RID: 346 RVA: 0x0000572C File Offset: 0x0000392C
		public bool IsPptShowTypeSpeaker
		{
			get
			{
				return base.GetSettingValue(ConfigController.IsPptShowTypeSpeakerKey, true);
			}
			set
			{
				base.SetSettingValue(ConfigController.IsPptShowTypeSpeakerKey, value);
			}
		}

		// Token: 0x04000064 RID: 100
		private static readonly string IsComOkKey = "IsComOk";

		// Token: 0x04000065 RID: 101
		private static readonly string IsShowingKey = "IsShowing";

		// Token: 0x04000066 RID: 102
		private static readonly string IsNeedStartKey = "IsNeedStart";

		// Token: 0x04000067 RID: 103
		private static readonly string OfficeVersionKey = "OfficeVersion";

		// Token: 0x04000068 RID: 104
		// Token: 0x04000069 RID: 105
		private static readonly string IsPptShowTypeSpeakerKey = "IsPptShowTypeSpeaker";

		/// <summary>
		/// Build the overlay window at start-up rather than at the first
		/// presentation. On by default: the entry animation needs a window that has
		/// already been presented, and paying for that once at start-up is cheaper
		/// than paying for it in the middle of show-begin.
		/// </summary>
		public bool IsPreload
		{
			get
			{
				return base.GetSettingValue(ConfigController.IsPreloadKey, true);
			}
			set
			{
				base.SetSettingValue(ConfigController.IsPreloadKey, value);
			}
		}

		/// <summary>
		/// Fetch every thumbnail as soon as a presentation starts, rather than
		/// only when the panel is first opened. On by default, so the panel comes
		/// up complete; off moves the pass to the first open, which starts the
		/// show with less work behind the entry animation.
		/// </summary>
		public bool IsPrefetch
		{
			get
			{
				return base.GetSettingValue(ConfigController.IsPrefetchKey, true);
			}
			set
			{
				base.SetSettingValue(ConfigController.IsPrefetchKey, value);
			}
		}

		// Token: 0x0400006A RID: 106
		private static readonly string IsInitApplicationSuccessfulKey = "IsInitApplicationSuccessful";

		private static readonly string IsPreloadKey = "IsPreload";

		private static readonly string IsPrefetchKey = "IsPrefetch";

		/// <summary>
		/// The colour mode of the interface: system / light / dark, as stored.
		/// system follows whatever Windows is set to, live; the other two pin it.
		/// </summary>
		public string ThemeMode
		{
			get
			{
				return base.GetSettingValue(ConfigController.ThemeModeKey, ToolbarTheme.ModeSystem);
			}
			set
			{
				base.SetSettingValue(ConfigController.ThemeModeKey, value);
			}
		}

		/// <summary>
		/// How wide thumbnails are decoded, in device-independent pixels; 0 means
		/// decode at the export's own size. This is the sharpness-versus-memory
		/// trade for the panel's list: the deck's images exist only in memory
		/// while a show is on, and every pixel of decode width above what the
		/// panel can display is memory with nothing to show for it. Chosen in the
		/// settings page; see MainWindow.LoadThumbnail for how it is applied.
		/// </summary>
		public int ThumbnailDecodeWidth
		{
			get
			{
				return base.GetSettingValue(ConfigController.ThumbnailDecodeWidthKey, ConfigController.DefaultThumbnailDecodeWidth);
			}
			set
			{
				base.SetSettingValue(ConfigController.ThumbnailDecodeWidthKey, value);
			}
		}

		/// <summary>
		/// The toolbar backdrop opacity of the default (hover) state, in whole
		/// percent. Chosen in the settings page; see ToolbarAppearance.Opacity.
		/// The old alpha was never stored a day in its life - every restart threw
		/// the slider's value away - so both states store like every other
		/// setting instead.
		/// </summary>
		public int ToolbarOpacityPercent
		{
			get
			{
				return base.GetSettingValue(ConfigController.ToolbarOpacityPercentKey, ConfigController.DefaultToolbarOpacityPercent);
			}
			set
			{
				base.SetSettingValue(ConfigController.ToolbarOpacityPercentKey, value);
			}
		}

		/// <summary>
		/// The toolbar backdrop opacity of the idle ("not clicked") state, in
		/// whole percent; see ToolbarAppearance.IdleOpacity.
		/// </summary>
		public int ToolbarIdleOpacityPercent
		{
			get
			{
				return base.GetSettingValue(ConfigController.ToolbarIdleOpacityPercentKey, ConfigController.DefaultToolbarIdleOpacityPercent);
			}
			set
			{
				base.SetSettingValue(ConfigController.ToolbarIdleOpacityPercentKey, value);
			}
		}

		/// <summary>
		/// The decode width used when nothing is stored yet: wide enough to stay
		/// crisp at the largest toolbar tier and at 150% display scaling, while
		/// still decoding a typical 1280x720 slide at roughly a tenth of its
		/// footprint.
		/// </summary>
		internal const int DefaultThumbnailDecodeWidth = 768;

		private static readonly string ThemeModeKey = "ThemeMode";

		private static readonly string ThumbnailDecodeWidthKey = "ThumbnailDecodeWidth";

		private static readonly string ToolbarOpacityPercentKey = "ToolbarOpacityPercent";

		private static readonly string ToolbarIdleOpacityPercentKey = "ToolbarIdleOpacityPercent";

		/// <summary>65, the percent form of ToolbarAppearance.DefaultOpacity.</summary>
		internal const int DefaultToolbarOpacityPercent = 65;

		/// <summary>35, the percent form of ToolbarAppearance.DefaultIdleOpacity.</summary>
		internal const int DefaultToolbarIdleOpacityPercent = 35;
	}
}
