using System;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using ModernWpf.Controls;
using ModernWpf.Controls.Primitives;
using System.Windows.Threading;
using ModernWpf;

namespace SimplePPSX
{
	/// <summary>
	/// Application entry point.
	///
	/// The app lives in the tray and runs as a single instance while watching the
	/// PowerPoint presentation events. The UI itself is not created here: it is built
	/// by ViewHelper when a presentation actually starts (or right away, when "preload
	/// the toolbars" is on), which is why the main window is a transparent overlay
	/// floating above the slide.
	/// </summary>
	public partial class App : Application
	{
		private TrayIconService _trayIcon;

		protected override void OnExit(ExitEventArgs e)
		{
			if (this._trayIcon != null)
			{
				this._trayIcon.Dispose();
				this._trayIcon = null;
			}
			base.OnExit(e);
		}

		protected override void OnStartup(StartupEventArgs e)
		{

			// One-time set-up, all of it before the first window or element exists:
			// 1) square corners: ModernWpf still hands out its stock 4/2 radii through
			//    the theme dictionaries that live on the elements themselves (the
			//    NumberBox spin popup), so patch those too, now
			ThemeRadiusFix.Hook();
			// 1.5) the colour mode: the stored system / light / dark choice is
			//    pushed onto ModernWpf before any window exists, and the toolbar
			//    brush set is stamped to match. Both are safe to redo later -
			//    the settings page writes them again on every change.
			ThemeManager.Current.ApplicationTheme = ToolbarTheme.ToApplicationTheme(ViewHelper.ThemeMode);
			ToolbarTheme.Apply();
			// 2) push the current toolbar opacity onto the chrome that exists already.
			//    Controls register their backdrop when they are constructed, so this
			//    only has to cover whatever was built before start-up (nothing, in the
			//    normal flow) - the registration in each control does the rest.
			ToolbarAppearance.Apply();
			// 3) give the late-bound COM layer somewhere to report. Every member read
			//    and write on the presentation app goes through ComLate, so this is
			//    the single place a WPS gap (a member WPS does not implement) shows
			//    up in the log, instead of being swallowed at each call site.
			ComLate.LogFailure = delegate(string message)
			{
				Log.WriteDebug(message);
			};

			// 4) the app lives in the tray and has no window of its own for most of its
			//    life: the toolbar overlay is built when a presentation starts and is
			//    hidden again when it ends. WPF's default is OnLastWindowClose, so the
			//    moment the user closes the settings window - which is often the ONLY
			//    window open - WPF decides the application should shut down, taking the
			//    tray icon and the presentation watcher with it.
			//
			//    OnExplicitShutdown says who ends the process: TrayIconService's Exit
			//    and App.PostErrorAndShutdown, and nothing else. A window closing is not
			//    a reason to quit.
			this.ShutdownMode = ShutdownMode.OnExplicitShutdown;

			base.OnStartup(e);

			// Single instance is decided BEFORE anything is built.
			//
			// It used to be decided after StartProcess(), so a second launch built the
			// tray icon, the COM watcher and every view before quietly calling
			// Environment.Exit(0). Besides being wasted work, it left no chance to tell
			// the user anything - which is what this dialog is for.
			//
			// The restart case is the exception and must still start everything: the
			// outgoing process may still hold the mutex when Restart.SimplePPSX spawns
			// its replacement, so the replacement legitimately finds it taken.
			bool isFirstInstance;
			App._mutex = new Mutex(true, App.AppMutex, out isFirstInstance);

			bool isRestart = e.Args.Any<string>() && e.Args.First<string>() == App.AppRestart;
			if (!isFirstInstance && !isRestart)
			{
				this.ShowAlreadyRunningDialog();
				return;
			}

			this._trayIcon = new TrayIconService();
			this.StartProcess();
		}

		/// <summary>
		/// Tells the user another copy is already running, then quits this one.
		///
		/// A ContentDialog needs an owner window, and this instance has none - the
		/// tray and the toolbar overlay are built later, and building them just to
		/// show a dialog is the waste this whole path exists to avoid. So a throwaway
		/// zero-size window is parked off-screen purely to own the dialog, and closed
		/// again straight after.
		///
		/// Shutdown(0) rather than Environment.Exit(0), because Shutdown closes the
		/// dispatcher properly: it lets the dialog's window tear down and the app's
		/// own shutdown path run. ShutdownMode is already OnExplicitShutdown, so this
		/// is the call that actually ends the process.
		/// </summary>
		private void ShowAlreadyRunningDialog()
		{
			Window owner = new Window
			{
				Width = 0,
				Height = 0,
				WindowStyle = WindowStyle.None,
				ShowInTaskbar = false,
				ShowActivated = false,
				// Off-screen rather than merely invisible: a zero-size window still
				// flashes a border on some Windows builds.
				Left = -32000,
				Top = -32000
			};

			try
			{
				// Shown AND activated, or the dialog opens behind nothing and is never
				// seen. ShowActivated = false is set on the window above because the
				// window must not steal focus at start-up, but this one has to be
				// foreground for the user to read it - so it is activated here and only
				// here.
				owner.Show();
				owner.Activate();

				ContentDialog dialog = new ContentDialog
				{
					Owner = owner,
					Title = "已经在运行",
					Content = "SimplePPSX 已经有一个实例在运行，无需重复启动。",
					CloseButtonText = "知道了"
				};

				// ModernWpfUI 0.9.6 has no synchronous Show, so the task it returns has
				// to be pumped to completion HERE. Just calling ShowAsync and carrying on
				// is what made this silently do nothing: the call returned a Task that
				// had not run yet, the finally block below closed the owner window in
				// the same tick, and the dialog was torn down before it ever rendered.
				//
				// Wait() would deadlock - the dialog needs this very dispatcher to keep
				// pumping - so a nested DispatcherFrame is pushed instead, and the frame
				// is released by the continuation when the dialog closes.
				System.Threading.Tasks.Task result = dialog.ShowAsync();
				DispatcherFrame frame = new DispatcherFrame();
				result.ContinueWith(
					delegate { frame.Continue = false; },
					System.Threading.Tasks.TaskContinuationOptions.ExecuteSynchronously,
					System.Threading.Tasks.TaskScheduler.FromCurrentSynchronizationContext());

				Dispatcher.PushFrame(frame);
			}
			catch (Exception ex)
			{
				// A dialog that cannot be shown must not leave a second copy running and
				// must not take the app down with it: the original behaviour was to exit
				// silently, so falling back to that is the safe outcome.
				Log.WriteDebug("ShowAlreadyRunningDialog: " + ex.GetType().Name + ": " + ex.Message);
			}
			finally
			{
				try
				{
					owner.Close();
				}
				catch (Exception)
				{
					// Nothing useful to do; the process is about to end anyway.
				}
			}

			this.Shutdown(0);
		}



		private void StartProcess()
		{
			AppPath.ClearTempFolder();
			ConfigHelper.Instance = new ConfigHelper();
			this._pptOperator = new PptOperator();
			PptOperator.Instance = this._pptOperator;
			this._pptOperator.Initialize();
			MouseHelper.Instance = new MouseHelper();
			new ViewHelper().Initialize();
			App.Version = App.GetAttribute<AssemblyFileVersionAttribute>().Version;
			this.CatchUnhandledException();
		}

		private void CatchUnhandledException()
		{
			base.DispatcherUnhandledException += delegate(object s, DispatcherUnhandledExceptionEventArgs e)
			{
				this.PostErrorAndShutdown(new UnhandledExceptionEventArgs(e, false));
			};
			AppDomain.CurrentDomain.UnhandledException += delegate(object s, UnhandledExceptionEventArgs e)
			{
				this.PostErrorAndShutdown(e);
			};
			Dispatcher.CurrentDispatcher.UnhandledException += delegate(object s, DispatcherUnhandledExceptionEventArgs e)
			{
				this.PostErrorAndShutdown(new UnhandledExceptionEventArgs(e, false));
			};
		}

		private void PostErrorAndShutdown(UnhandledExceptionEventArgs args)
		{
			Log.WriteError(DescribeException(args.ExceptionObject));
			if (ConfigHelper.Instance.IsComOk)
			{
				Restart.SimplePPSX();
			}
			Environment.Exit(0);
		}

		/// <summary>
		/// Renders whatever the exception plumbing handed us into something readable.
		///
		/// This is not cosmetic. Both dispatcher handlers wrap their exception in an
		/// <see cref="UnhandledExceptionEventArgs"/> and pass that straight through, so
		/// args.ExceptionObject is the DispatcherUnhandledExceptionEventArgs WRAPPER, not
		/// the exception inside it. Calling ToString() on the wrapper writes the type
		/// name and nothing else - which is exactly what error.txt contained:
		///
		///     2026/10/6 13:14:52  System.Windows.Threading.DispatcherUnhandledExceptionEventArgs
		///
		/// repeated once per crash, with no message, no type and no stack. Every
		/// unhandled exception in this app was therefore indistinguishable and
		/// untraceable, which is why a window that "does not open" could not be
		/// diagnosed from the log at all.
		/// </summary>
		private static string DescribeException(object exceptionObject)
		{
			Exception real = Unwrap(exceptionObject);
			if (real == null)
			{
				return "unknown failure: " + (exceptionObject == null ? "(null)" : exceptionObject.ToString());
			}

			// AggregateException is what the dispatcher produces when several failures
			// land at once; reporting only InnerException would hide all but one.
			AggregateException aggregate = real as AggregateException;
			if (aggregate != null && aggregate.InnerExceptions.Count > 1)
			{
				System.Text.StringBuilder all = new System.Text.StringBuilder();
				all.AppendLine(real.GetType().FullName + " with " + aggregate.InnerExceptions.Count + " inner exceptions:");
				foreach (Exception inner in aggregate.InnerExceptions)
				{
					all.AppendLine(inner.ToString());
					all.AppendLine("----");
				}
				return all.ToString();
			}

			return real.ToString();
		}

		/// <summary>
		/// Digs the real Exception out of the wrappers the two paths above build.
		/// </summary>
		private static Exception Unwrap(object exceptionObject)
		{
			for (int depth = 0; depth < 8 && exceptionObject != null; depth++)
			{
				if (exceptionObject is Exception direct)
				{
					return direct;
				}

				// DispatcherUnhandledExceptionEventArgs.Exception
				object next = exceptionObject.GetType()
					.GetProperty("Exception", BindingFlags.Instance | BindingFlags.Public)
					?.GetValue(exceptionObject, null);

				if (next == null)
				{
					// UnhandledExceptionEventArgs.ExceptionObject
					next = exceptionObject.GetType()
						.GetProperty("ExceptionObject", BindingFlags.Instance | BindingFlags.Public)
						?.GetValue(exceptionObject, null);
				}

				exceptionObject = next;
			}

			return null;
		}

		private static T GetAttribute<T>() where T : Attribute
		{
			return (T)((object)Assembly.GetAssembly(Application.Current.GetType()).GetCustomAttributes(typeof(T), false)[0]);
		}

		public static string Version;

		private static string AppRestart = "Restart";

		public static readonly string AppName = "SimplePPSX";

		private static string AppMutex = "#P_#P_#T_&S_&e_&r_*v_*i_*c_$e";

		private PptOperator _pptOperator;

		private static Mutex _mutex;
	}
}
