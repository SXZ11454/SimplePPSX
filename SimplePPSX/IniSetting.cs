using System;
using System.Globalization;

namespace SimplePPSX
{
	// Token: 0x02000026 RID: 38
	public abstract class IniSetting
	{
		// Token: 0x06000137 RID: 311 RVA: 0x00005402 File Offset: 0x00003602
		protected IniSetting(string settingPath, string section)
		{
			this._settingPath = settingPath;
			this._section = section;
		}

		// Token: 0x06000138 RID: 312 RVA: 0x00005418 File Offset: 0x00003618
		public T GetSettingValue<T>(string key, T defaultValue, Func<string, T, T> convertFunc, string section, string path)
		{
			string text = IniUtility.Read(section, key, this._settingPath);
			if (!string.IsNullOrEmpty(text))
			{
				return convertFunc(text, defaultValue);
			}
			return defaultValue;
		}

		// Token: 0x06000139 RID: 313 RVA: 0x00005446 File Offset: 0x00003646
		public string GetSettingValue(string key, string defaultValue, string section, string path)
		{
			return this.GetSettingValue<string>(key, defaultValue, delegate(string valueToConvert, string defaultValueToConvert)
			{
				if (!string.IsNullOrEmpty(valueToConvert))
				{
					return valueToConvert;
				}
				return defaultValueToConvert;
			}, section, path);
		}

		// Token: 0x0600013A RID: 314 RVA: 0x00005472 File Offset: 0x00003672
		public string GetSettingValue(string key, string defaultValue, string section)
		{
			return this.GetSettingValue(key, defaultValue, section, this._settingPath);
		}

		// Token: 0x0600013B RID: 315 RVA: 0x00005483 File Offset: 0x00003683
		public string GetSettingValue(string key, string defaultValue)
		{
			return this.GetSettingValue(key, defaultValue, this._section, this._settingPath);
		}

		// Token: 0x0600013C RID: 316 RVA: 0x00005499 File Offset: 0x00003699
		public int GetSettingValue(string key, int defaultValue, string section, string path)
		{
			return this.GetSettingValue<int>(key, defaultValue, delegate(string valueToConvert, int defaultValueToConvert)
			{
				int num;
				if (!int.TryParse(valueToConvert, out num))
				{
					return defaultValueToConvert;
				}
				return num;
			}, section, path);
		}

		// Token: 0x0600013D RID: 317 RVA: 0x000054C5 File Offset: 0x000036C5
		public int GetSettingValue(string key, int defaultValue, string section)
		{
			return this.GetSettingValue(key, defaultValue, section, this._settingPath);
		}

		// Token: 0x0600013E RID: 318 RVA: 0x000054D6 File Offset: 0x000036D6
		public int GetSettingValue(string key, int defaultValue)
		{
			return this.GetSettingValue(key, defaultValue, this._section, this._settingPath);
		}

		// Token: 0x0600013F RID: 319 RVA: 0x000054EC File Offset: 0x000036EC
		public bool GetSettingValue(string key, bool defaultValue, string section, string path)
		{
			return this.GetSettingValue<bool>(key, defaultValue, delegate(string valueToConvert, bool defaultValueToConvert)
			{
				int num;
				if (!int.TryParse(valueToConvert, out num))
				{
					return defaultValueToConvert;
				}
				return 1 == num;
			}, section, path);
		}

		// Token: 0x06000140 RID: 320 RVA: 0x00005518 File Offset: 0x00003718
		public bool GetSettingValue(string key, bool defaultValue, string section)
		{
			return this.GetSettingValue(key, defaultValue, section, this._settingPath);
		}

		// Token: 0x06000141 RID: 321 RVA: 0x00005529 File Offset: 0x00003729
		public bool GetSettingValue(string key, bool defaultValue)
		{
			return this.GetSettingValue(key, defaultValue, this._section, this._settingPath);
		}

		// Token: 0x06000142 RID: 322 RVA: 0x0000553F File Offset: 0x0000373F
		public void SetSettingValue<T>(string key, T value, Func<T, string> convertFunc, string section, string path)
		{
			IniUtility.Write(section, key, convertFunc(value), this._settingPath);
		}

		// Token: 0x06000143 RID: 323 RVA: 0x00005556 File Offset: 0x00003756
		public void SetSettingValue(string key, string value, string section, string path)
		{
			this.SetSettingValue<string>(key, value, (string valueToConvert) => valueToConvert, section, path);
		}

		// Token: 0x06000144 RID: 324 RVA: 0x00005582 File Offset: 0x00003782
		public void SetSettingValue(string key, string value, string section)
		{
			this.SetSettingValue(key, value, section, this._settingPath);
		}

		// Token: 0x06000145 RID: 325 RVA: 0x00005593 File Offset: 0x00003793
		public void SetSettingValue(string key, string value)
		{
			this.SetSettingValue(key, value, this._section, this._settingPath);
		}

		// Token: 0x06000146 RID: 326 RVA: 0x000055A9 File Offset: 0x000037A9
		public void SetSettingValue(string key, int value, string section, string path)
		{
			this.SetSettingValue<int>(key, value, (int valueToConvert) => valueToConvert.ToString(CultureInfo.InvariantCulture), section, path);
		}

		// Token: 0x06000147 RID: 327 RVA: 0x000055D5 File Offset: 0x000037D5
		public void SetSettingValue(string key, int value, string section)
		{
			this.SetSettingValue(key, value, section, this._settingPath);
		}

		// Token: 0x06000148 RID: 328 RVA: 0x000055E6 File Offset: 0x000037E6
		public void SetSettingValue(string key, int value)
		{
			this.SetSettingValue(key, value, this._section, this._settingPath);
		}

		// Token: 0x06000149 RID: 329 RVA: 0x000055FC File Offset: 0x000037FC
		public void SetSettingValue(string key, bool value, string section, string path)
		{
			this.SetSettingValue<bool>(key, value, delegate(bool valueToConvert)
			{
				if (!valueToConvert)
				{
					return "0";
				}
				return "1";
			}, section, path);
		}

		// Token: 0x0600014A RID: 330 RVA: 0x00005628 File Offset: 0x00003828
		public void SetSettingValue(string key, bool value, string section)
		{
			this.SetSettingValue(key, value, section, this._settingPath);
		}

		// Token: 0x0600014B RID: 331 RVA: 0x00005639 File Offset: 0x00003839
		public void SetSettingValue(string key, bool value)
		{
			this.SetSettingValue(key, value, this._section, this._settingPath);
		}

		// Token: 0x04000062 RID: 98
		private readonly string _section;

		// Token: 0x04000063 RID: 99
		private readonly string _settingPath;
	}
}
