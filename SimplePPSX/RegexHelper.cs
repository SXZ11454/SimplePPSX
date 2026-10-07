using System;
using System.Text.RegularExpressions;

namespace SimplePPSX
{
	// Token: 0x0200001C RID: 28
	public static class RegexHelper
	{
		// Token: 0x060000EF RID: 239 RVA: 0x00004816 File Offset: 0x00002A16
		public static bool IsLocalPath(string path)
		{
			return new Regex("^(?<fpath>([a-zA-Z]:\\\\)([\\s\\.\\-\\w]+\\\\)*)(?<fname>[\\w]+.[\\w]+)").Match(path).Success;
		}
	}
}
