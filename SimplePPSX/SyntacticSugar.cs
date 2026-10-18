using System;
using System.Windows;
using System.Windows.Threading;

namespace SimplePPSX
{
	// Token: 0x0200003B RID: 59
	public static class SyntacticSugar
	{
		// Token: 0x0600026A RID: 618 RVA: 0x0000B048 File Offset: 0x00009248
		public static void TryCatchAction(Action tryAction, Action finallyAction = null)
		{
			try
			{
				tryAction();
			}
			catch (Exception ex)
			{
				Log.WriteError(ex.ToString());
			}
			finally
			{
				if (finallyAction != null)
				{
					finallyAction();
				}
			}
		}

		// Token: 0x0600026B RID: 619 RVA: 0x0000B090 File Offset: 0x00009290
		public static void CurrentDispatcherInvoke(Action action)
		{
			Application.Current.Dispatcher.Invoke(new Action(action.Invoke));
		}

		// Token: 0x0600026C RID: 620 RVA: 0x0000B0AD File Offset: 0x000092AD
		public static void CurrentDispatcherBeginInvoke(Action action)
		{
			Application.Current.Dispatcher.BeginInvoke(new Action(action.Invoke), new object[0]);
		}

		// Token: 0x0600026D RID: 621 RVA: 0x0000B0D1 File Offset: 0x000092D1
		public static void CurrentDispatcherBeginInvoke(DispatcherPriority dispatcherPriority, Action action)
		{
			Application.Current.Dispatcher.BeginInvoke(dispatcherPriority, new Action(action.Invoke));
		}

		/// <summary>
		/// Runs the action on the dispatcher thread: inline when the caller already is
		/// that thread, posted otherwise.
		///
		/// CurrentDispatcherInvoke waits - and the threads that reach for this are
		/// PowerPoint's own event callbacks, which must not wait on the UI thread.
		/// While such a callback waits, PowerPoint is not pumping, and the UI thread
		/// (a COM call of its own at that exact moment, say) waits for PowerPoint in
		/// turn: both sides parked on each other until COM gives up. That crossing is
		/// what used to freeze the overlay for seconds. Posting instead lets the event
		/// thread go straight back to PowerPoint's pump. Order is kept by the
		/// dispatcher queue - same priority, posted in call order.
		/// </summary>
		public static void DispatcherPostOrInline(Action action)
		{
			Dispatcher dispatcher = Application.Current.Dispatcher;
			if (dispatcher.CheckAccess())
			{
				action.Invoke();
				return;
			}
			dispatcher.BeginInvoke(new Action(action.Invoke), new object[0]);
		}
	}
}
