using System;
using System.Windows.Threading;

namespace SimplePPSX
{
	// Token: 0x02000018 RID: 24
	public class MouseHelper
	{
		// Token: 0x17000021 RID: 33
		// (get) Token: 0x060000D4 RID: 212 RVA: 0x000043D8 File Offset: 0x000025D8
		// (set) Token: 0x060000D5 RID: 213 RVA: 0x000043DF File Offset: 0x000025DF
		public static MouseHelper Instance { get; set; }

		// Token: 0x060000D6 RID: 214 RVA: 0x000043E7 File Offset: 0x000025E7
		public MouseHelper()
		{
			this._timer = new DispatcherTimer
			{
				Interval = TimeSpan.FromMilliseconds((double)MouseHelper.TimerSpan)
			};
			this._timer.Tick += this._timerTick;
		}

		// Token: 0x060000D7 RID: 215 RVA: 0x00004422 File Offset: 0x00002622
		public void Start()
		{
			if (MouseHelper._isStart)
			{
				return;
			}
			MouseHelper._isStart = true;
			this._timer.Start();
		}

		// Token: 0x060000D8 RID: 216 RVA: 0x0000443D File Offset: 0x0000263D
		public void Stop()
		{
			if (!MouseHelper._isStart)
			{
				return;
			}
			MouseHelper._isStart = false;
			this._timer.Stop();
		}

		// Token: 0x060000D9 RID: 217 RVA: 0x00004458 File Offset: 0x00002658
		private void _timerTick(object sender, EventArgs e)
		{
			MousePoint mousePoint;
			WindowsApi.GetCursorPos(out mousePoint);
			int num = Math.Abs(WindowsApi.GetKeyState(1));
			if ((double)num > MouseHelper.MouseStateBoundary && (double)this._preMouseState < MouseHelper.MouseStateBoundary)
			{
				MouseHelper.StampInput();
				this.OnDown(new MouseHelperEventArgs((double)mousePoint.X, (double)mousePoint.Y));
				this._preMouseState = num;
			}
			if ((double)num > MouseHelper.MouseStateBoundary && (double)this._preMouseState > MouseHelper.MouseStateBoundary)
			{
				MouseHelper.StampInput();
				this.OnMove(new MouseHelperEventArgs((double)mousePoint.X, (double)mousePoint.Y));
				this._preMouseState = num;
			}
			if ((double)num < MouseHelper.MouseStateBoundary && (double)this._preMouseState > MouseHelper.MouseStateBoundary)
			{
				MouseHelper.StampInput();
				this.OnUp(new MouseHelperEventArgs((double)mousePoint.X, (double)mousePoint.Y));
				this._preMouseState = num;
			}
		}

		// Token: 0x14000004 RID: 4
		// (add) Token: 0x060000DA RID: 218 RVA: 0x00004520 File Offset: 0x00002720
		// (remove) Token: 0x060000DB RID: 219 RVA: 0x00004558 File Offset: 0x00002758
		public event EventHandler<MouseHelperEventArgs> Down;

		// Token: 0x14000005 RID: 5
		// (add) Token: 0x060000DC RID: 220 RVA: 0x00004590 File Offset: 0x00002790
		// (remove) Token: 0x060000DD RID: 221 RVA: 0x000045C8 File Offset: 0x000027C8
		public event EventHandler<MouseHelperEventArgs> Move;

		// Token: 0x14000006 RID: 6
		// (add) Token: 0x060000DE RID: 222 RVA: 0x00004600 File Offset: 0x00002800
		// (remove) Token: 0x060000DF RID: 223 RVA: 0x00004638 File Offset: 0x00002838
		public event EventHandler<MouseHelperEventArgs> Up;

		// Token: 0x060000E0 RID: 224 RVA: 0x0000466D File Offset: 0x0000286D
		protected virtual void OnDown(MouseHelperEventArgs e)
		{
			EventHandler<MouseHelperEventArgs> down = this.Down;
			if (down == null)
			{
				return;
			}
			down(this, e);
		}

		// Token: 0x060000E1 RID: 225 RVA: 0x00004681 File Offset: 0x00002881
		protected virtual void OnMove(MouseHelperEventArgs e)
		{
			EventHandler<MouseHelperEventArgs> move = this.Move;
			if (move == null)
			{
				return;
			}
			move(this, e);
		}

		// Token: 0x060000E2 RID: 226 RVA: 0x00004695 File Offset: 0x00002895
		protected virtual void OnUp(MouseHelperEventArgs e)
		{
			EventHandler<MouseHelperEventArgs> up = this.Up;
			if (up == null)
			{
				return;
			}
			up(this, e);
		}

		// Token: 0x04000040 RID: 64
		private static bool _isStart;

		// Token: 0x04000041 RID: 65
		private int _preMouseState;

		// Token: 0x04000042 RID: 66
		private readonly DispatcherTimer _timer;

		// Token: 0x04000043 RID: 67
		private static readonly int TimerSpan = 6;

		// Token: 0x04000044 RID: 68
		private static readonly double MouseStateBoundary = 25.0;

		/// <summary>TickCount64 of the last observed press/drag/release.</summary>
		private static long _lastInputTick;

		/// <summary>
		/// Stamps the moment input was last seen. The export background reads
		/// this: a thumbnail render that starts while the user is mid-click is
		/// a full offscreen slide render on PowerPoint's STA - the same thread
		/// that has to answer the click - and that race is exactly the "clicked
		/// but nothing happened" feel. The stamp costs one store per observed
		/// press or drag, inside the 6ms poll that already exists.
		/// </summary>
		internal static void StampInput()
		{
			System.Threading.Interlocked.Exchange(ref MouseHelper._lastInputTick, System.Environment.TickCount64);
		}

		/// <summary>Milliseconds since the last observed press/drag/release.</summary>
		internal static long MillisecondsSinceInput
		{
			get
			{
				return System.Environment.TickCount64 - System.Threading.Interlocked.Read(ref MouseHelper._lastInputTick);
			}
		}
	}
}
