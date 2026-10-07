using System;
using System.Runtime.InteropServices;

namespace SimplePPSX
{
	/// <summary>
	/// Small COM helpers that .NET Framework offered out of the box and .NET 6/10
	/// do not.
	///
	/// <see cref="Marshal.GetActiveObject(string)"/> was removed: .NET Core has no
	/// managed entry point into the Running Object Table, but oleaut32.dll still
	/// exports one, so it is declared here. That call is how the app attaches to the
	/// PowerPoint instance that is already running, which is the whole point of
	/// PptOperator.CreatePowerPointApplication.
	/// </summary>
	internal static class ComObject
	{
		[DllImport("ole32.dll", CharSet = CharSet.Unicode, ExactSpelling = true, PreserveSig = false)]
		private static extern void CLSIDFromProgIDEx([MarshalAs(UnmanagedType.LPWStr)] string progId, out Guid clsid);

		[DllImport("oleaut32.dll", CharSet = CharSet.Unicode, ExactSpelling = true, PreserveSig = false)]
		private static extern void GetActiveObject(ref Guid clsid, IntPtr reserved, [MarshalAs(UnmanagedType.IUnknown)] out object obj);

		/// <summary>
		/// Returns the running COM object registered under the given ProgID, or null
		/// when nothing is registered (the caller usually then starts the program
		/// itself and retries).
		/// </summary>
		internal static object TryGetActiveObject(string progId)
		{
			if (string.IsNullOrEmpty(progId))
			{
				return null;
			}

			Guid clsid;
			try
			{
				CLSIDFromProgIDEx(progId, out clsid);
			}
			catch (Exception)
			{
				// ProgID not registered at all
				return null;
			}

			try
			{
				object instance;
				GetActiveObject(ref clsid, IntPtr.Zero, out instance);
				return instance;
			}
			catch (COMException)
			{
				// registered but not running (RPC_E_CLASSNOTAVAILABLE)
				return null;
			}
			catch (Exception)
			{
				return null;
			}
		}
	}
}
