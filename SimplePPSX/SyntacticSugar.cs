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
	}
}
