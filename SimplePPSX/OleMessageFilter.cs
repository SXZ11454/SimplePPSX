using System;
using System.Runtime.InteropServices;

namespace SimplePPSX
{
	/// <summary>
	/// A client-side OLE message filter for the UI thread's STA.
	///
	/// WHY THIS EXISTS
	/// ===============
	/// Every now and then the UI thread has to cross into the presentation
	/// application's STA - a pointer write, a geometry read that missed the
	/// cache. When that server is busy (a deck animating; the thumbnail export
	/// in full swing) its message filter rejects the call with
	/// RPC_E_SERVERCALL_RETRYLATER, and Windows' own built-in filter then
	/// quietly retries the call ON THE CALLING THREAD for up to about ten
	/// seconds. On a worker thread nobody ever notices. On the UI thread that
	/// is the overlay freezing between clicks.
	///
	/// Registering this filter on the UI STA changes the one knob that matters:
	/// <see cref="Filter.RetryRejectedCall"/> returns -1, which cancels the call
	/// immediately - the caller receives RPC_E_CALL_REJECTED at once and every
	/// ComLate call site already falls back to its last known value. No click
	/// can ever be parked for ten seconds waiting on a busy presentation again.
	///
	/// Installed once, at start-up, on the main thread. Other STAs (there are
	/// none today) would need their own registration; the background MTA threads
	/// never consult a client filter, so the export keeps its patient retries.
	/// </summary>
	public static class OleMessageFilter
	{
		/// <summary>
		/// Registers the filter on this thread. Idempotent. The thread must be
		/// an STA for the registration to have any meaning; the app's dispatcher
		/// thread is one.
		/// </summary>
		public static void Install()
		{
			if (OleMessageFilter.Installed)
			{
				return;
			}
			try
			{
				int result = OleMessageFilter.CoRegisterMessageFilter(new OleMessageFilter.Filter(), out OleMessageFilter.IMessageFilter previous);
				OleMessageFilter.Installed = (result == 0);
			}
			catch (Exception installationFailure)
			{
				// A missing registration is the old ten-second behaviour, not a
				// crash: report it and carry on.
				Log.WriteError("OleMessageFilter.Install: " + installationFailure.Message);
			}
		}

		private static bool Installed;

		[DllImport("ole32.dll")]
		private static extern int CoRegisterMessageFilter(OleMessageFilter.IMessageFilter newFilter, out OleMessageFilter.IMessageFilter previousFilter);

		/// <summary>
		/// IOleMessageFilter verbatim (00000016-...). PreserveSig everywhere:
		/// the return values are themselves COM codes the OLE layer inspects.
		/// </summary>
		[ComImport]
		[Guid("00000016-0000-0000-C000-000000000046")]
		[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
		private interface IMessageFilter
		{
			[PreserveSig]
			int HandleInComingCall(int callType, IntPtr callerTask, int tickCount, IntPtr interfaceInfo);

			[PreserveSig]
			int RetryRejectedCall(IntPtr calleeTask, int tickCount, int rejectedType);

			[PreserveSig]
			int MessagePending(IntPtr calleeTask, int tickCount, int pendingType);
		}

		private class Filter : OleMessageFilter.IMessageFilter
		{
			/// <summary>Incoming calls are handled as usual (SERVERCALL_ISHANDLED).</summary>
			public int HandleInComingCall(int callType, IntPtr callerTask, int tickCount, IntPtr interfaceInfo)
			{
				return 0;
			}

			/// <summary>
			/// A rejected outbound call is cancelled outright (-1). Never queue a
			/// wait: the caller falls back, and the next poll cycle asks again if
			/// the answer still matters.
			/// </summary>
			public int RetryRejectedCall(IntPtr calleeTask, int tickCount, int rejectedType)
			{
				return -1;
			}

			/// <summary>Let the default pending-message handling stand (2).</summary>
			public int MessagePending(IntPtr calleeTask, int tickCount, int pendingType)
			{
				return 2;
			}
		}
	}
}
