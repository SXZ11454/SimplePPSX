using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace SimplePPSX
{
	// Token: 0x02000033 RID: 51
	public class Md5SecurityUtility
	{
		// Token: 0x060001CF RID: 463 RVA: 0x00007D7C File Offset: 0x00005F7C
		public static string GetShaFromFile(string filePath)
		{
			if (!File.Exists(filePath))
			{
				return null;
			}
			string text;
			try
			{
				using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
				{
					// SHA256Cng was removed after .NET Framework; SHA256.Create() is the
					// portable way and returns the very same algorithm on Windows.
					using (SHA256 sha256 = SHA256.Create())
					{
						byte[] array = sha256.ComputeHash(fileStream);
						StringBuilder stringBuilder = new StringBuilder();
						for (int i = 0; i < array.Length; i++)
						{
							stringBuilder.Append(array[i].ToString("x2"));
						}
						text = stringBuilder.ToString();
					}
				}
			}
			catch (Exception ex)
			{
				Log.WriteError(ex.ToString());
				text = null;
			}
			return text;
		}
	}
}
