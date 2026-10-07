using System;
using System.Runtime.InteropServices;
using System.Text;

namespace SimplePPSX
{
	// Token: 0x02000022 RID: 34
	public static class IniUtility
	{
		// Token: 0x06000110 RID: 272
		[DllImport("kernel32.dll", CharSet = CharSet.Ansi)]
		public static extern long WritePrivateProfileString(string section, string key, string val, string filePath);

		// Token: 0x06000111 RID: 273
		[DllImport("kernel32.dll", CharSet = CharSet.Ansi)]
		public static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filePath);

		// Token: 0x06000112 RID: 274 RVA: 0x00004DF4 File Offset: 0x00002FF4
		public static void Write(string section, string key, string value, string filePath)
		{
			IniUtility.WritePrivateProfileString(section, key, value, filePath);
		}

		// Token: 0x06000113 RID: 275 RVA: 0x00004E00 File Offset: 0x00003000
		public static string Read(string section, string key, string filePath)
		{
			StringBuilder stringBuilder = new StringBuilder(1024);
			IniUtility.GetPrivateProfileString(section, key, "", stringBuilder, 1024, filePath);
			return stringBuilder.ToString();
		}
	}
}
