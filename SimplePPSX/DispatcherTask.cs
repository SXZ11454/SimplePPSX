using System;
using System.Windows.Threading;

namespace SimplePPSX
{
	// Token: 0x02000014 RID: 20
	public class DispatcherTask
	{
		// Token: 0x060000B6 RID: 182 RVA: 0x00003E2D File Offset: 0x0000202D
		public static DispatcherTask Start(Action action)
		{
			return DispatcherTask.Start(0.0, action);
		}

		// Token: 0x060000B7 RID: 183 RVA: 0x00003E3E File Offset: 0x0000203E
		public static DispatcherTask Start(double milliseconds, Action action)
		{
			DispatcherTask dispatcherTask = new DispatcherTask(milliseconds, action);
			dispatcherTask.Start();
			return dispatcherTask;
		}

		// Token: 0x060000B8 RID: 184 RVA: 0x00003E4D File Offset: 0x0000204D
		public static DispatcherTask StartLoop(int loopCount, double intervalMilliseconds, Action action)
		{
			DispatcherTask dispatcherTask = new DispatcherTask(intervalMilliseconds, action, loopCount);
			dispatcherTask.Start();
			return dispatcherTask;
		}

		// Token: 0x060000B9 RID: 185 RVA: 0x00003E5D File Offset: 0x0000205D
		public static DispatcherTask StartInfiniteLoop(double intervalMilliseconds, Action action)
		{
			DispatcherTask dispatcherTask = new DispatcherTask(intervalMilliseconds, action, -1);
			dispatcherTask.Start();
			return dispatcherTask;
		}

		// Token: 0x1700001F RID: 31
		// (get) Token: 0x060000BA RID: 186 RVA: 0x00003E6D File Offset: 0x0000206D
		// (set) Token: 0x060000BB RID: 187 RVA: 0x00003E75 File Offset: 0x00002075
		public bool HasStarted { get; private set; }

		// Token: 0x17000020 RID: 32
		// (get) Token: 0x060000BC RID: 188 RVA: 0x00003E7E File Offset: 0x0000207E
		// (set) Token: 0x060000BD RID: 189 RVA: 0x00003E86 File Offset: 0x00002086
		public bool SwallowsException { get; set; }

		// Token: 0x060000BE RID: 190 RVA: 0x00003E8F File Offset: 0x0000208F
		public void Start()
		{
			if (this.HasStarted)
			{
				throw new InvalidOperationException("This task has started already.");
			}
			this._timer.Start();
			this.HasStarted = true;
		}

		// Token: 0x060000BF RID: 191 RVA: 0x00003EB6 File Offset: 0x000020B6
		public void Stop()
		{
			this._timer.Stop();
			this.HasStarted = false;
		}

		// Token: 0x060000C0 RID: 192 RVA: 0x00003ECA File Offset: 0x000020CA
		public void Restart()
		{
			this._hitCount = 0;
			this._timer.Stop();
			this._timer.Start();
			this.HasStarted = true;
		}

		// Token: 0x060000C1 RID: 193 RVA: 0x00003EF0 File Offset: 0x000020F0
		public DispatcherTask(Action action)
			: this(0.0, action, 1)
		{
		}

		// Token: 0x060000C2 RID: 194 RVA: 0x00003F03 File Offset: 0x00002103
		public DispatcherTask(double milliseconds, Action action)
			: this(milliseconds, action, 1)
		{
		}

		// Token: 0x060000C3 RID: 195 RVA: 0x00003F10 File Offset: 0x00002110
		public DispatcherTask(double milliseconds, Action action, int loopCount)
		{
			if (action == null)
			{
				throw new ArgumentNullException("action");
			}
			if (milliseconds < 0.0)
			{
				throw new ArgumentOutOfRangeException("milliseconds");
			}
			this._action = action;
			this._timer = new DispatcherTimer
			{
				Interval = TimeSpan.FromMilliseconds(milliseconds)
			};
			this._timer.Tick += this.OnTimerTick;
			this._loopCount = loopCount;
		}

		// Token: 0x060000C4 RID: 196 RVA: 0x00003F84 File Offset: 0x00002184
		private void OnTimerTick(object sender, EventArgs e)
		{
			this._hitCount++;
			DispatcherTimer dispatcherTimer = (DispatcherTimer)sender;
			if (this._hitCount == this._loopCount)
			{
				dispatcherTimer.Stop();
			}
			if (this.SwallowsException)
			{
				try
				{
					this._action();
					return;
				}
				catch (Exception)
				{
					return;
				}
			}
			this._action();
		}

		// Token: 0x04000032 RID: 50
		private readonly Action _action;

		// Token: 0x04000033 RID: 51
		private readonly DispatcherTimer _timer;

		// Token: 0x04000034 RID: 52
		private readonly int _loopCount;

		// Token: 0x04000035 RID: 53
		private int _hitCount;
	}
}
