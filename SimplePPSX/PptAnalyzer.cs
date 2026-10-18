using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using Microsoft.Office.Interop.PowerPoint;

namespace SimplePPSX
{
	// Token: 0x02000036 RID: 54
	public static class PptAnalyzer
	{
		// Token: 0x17000044 RID: 68
		// (get) Token: 0x060001FC RID: 508 RVA: 0x000095C0 File Offset: 0x000077C0
		public static bool IsAnalyzeAllAsyncRunning
		{
			get
			{
				return PptAnalyzer._thread != null && PptAnalyzer._thread.IsAlive;
			}
		}

		// Token: 0x060001FD RID: 509 RVA: 0x000095D8 File Offset: 0x000077D8
		public static void StartAnalyzeAllAsync(AnalyzeRate rate)
		{
			if (PptAnalyzer.IsAnalyzeOver() || PptAnalyzer._isAnalyzingAll)
			{
				return;
			}
			if (!PptAnalyzer.IsAnalyzeAllAsyncRunning && PptOperator.Instance.IsShowing)
			{
				PptAnalyzer._thread = new Thread(new ThreadStart(PptAnalyzer.AnalyzeAsync))
				{
					IsBackground = true,
					// BelowNormal: this thread competes with the UI thread and
					// with PowerPoint itself for cores while the show is on
					// screen; the thumbnails can wait a frame, the show cannot.
					Priority = ThreadPriority.BelowNormal
				};
				PptAnalyzer._isBreakExportAsync = false;
				PptAnalyzer._thread.Start();
			}
			else if (!PptAnalyzer.IsAnalyzeAllAsyncRunning)
			{
				// The thread only starts from inside a show; say so when it was
				// asked for from anywhere else, rather than declining in silence.
				Log.WriteDebug("StartAnalyzeAllAsync: no thread started (show not running)");
			}
			PptAnalyzer.CurrentAnalyzeRate = rate;
			switch (PptAnalyzer.CurrentAnalyzeRate)
			{
			case AnalyzeRate.Lowest:
				PptAnalyzer._waitTime = 999;
				return;
			case AnalyzeRate.Middle:
				PptAnalyzer._waitTime = 333;
				return;
			case AnalyzeRate.Highest:
				PptAnalyzer._waitTime = 0;
				return;
			default:
				return;
			}
		}

		/// <summary>
		/// One-line state of the analysis thread, for the freeze watchdog.
		/// alive/BelowNormal with a sensible analyzedCount is the export at work;
		/// dead with the count parked below the deck means it stopped early.
		/// </summary>
		internal static string DescribeAnalyzer()
		{
			Thread thread = PptAnalyzer._thread;
			string state = (thread == null) ? "none" : (thread.IsAlive ? ("alive/" + thread.Priority) : "dead");
			try
			{
				return "thread=" + state + ", analyzedCount=" + PptAnalyzer.AnalyzedIndexAndThumbnailFilePathDictionary.Count;
			}
			catch (Exception)
			{
				return "thread=" + state + ", analyzedCount=? ";
			}
		}

		/// <summary>Set while the between-shows export pass is on the road.</summary>
		public static bool IsIdleExportRunning;

		/// <summary>
		/// Finishes the deck's thumbnail export while PowerPoint is idle.
		///
		/// Every thumbnail has to be rendered by PowerPoint itself, and a Slide
		/// Export is a full offscreen render on PowerPoint's STA - the same
		/// thread that runs the live show. Spread over a show, those renders
		/// are the stutter and the "stuck" seconds of a panel filling; run
		/// while no show is on screen they are cheap and invisible. So the
		/// pass is parked at the end of a show: whatever the deck is still
		/// missing is exported then, at an easy pace, and the next
		/// presentation finds the panel complete with no export of its own.
		/// The loop stands down the moment a show starts (IsShowing) or a
		/// stop is asked for.
		/// </summary>
		public static void StartIdleExport()
		{
			try
			{
				if (PptAnalyzer.IsIdleExportRunning || PptAnalyzer._isAnalyzingAll || PptAnalyzer.IsAnalyzeAllAsyncRunning)
				{
					return;
				}
				if (PptAnalyzer.IsAnalyzeOver() || PptOperator.Instance == null || PptOperator.Instance.Slides == null || PptOperator.Instance.Slides.Count == 0 || PptAnalyzer.AnalyzedIndexAndThumbnailFilePathDictionary == null || PptAnalyzer.AnalyzedIndexAndThumbnailFilePathDictionary.IsEmpty || string.IsNullOrEmpty(PptAnalyzer.ThumbnailFilePath))
				{
					return;
				}
				// The show that just ended asked the stop flag on its way out;
				// this pass is the one that owns the flag now.
				PptAnalyzer._isBreakExportAsync = false;
				PptAnalyzer.IsIdleExportRunning = true;
				Thread thread = new Thread(new ThreadStart(PptAnalyzer.IdleExportLoop))
				{
					IsBackground = true,
					Priority = ThreadPriority.BelowNormal
				};
				thread.Start();
			}
			catch (Exception)
			{
				PptAnalyzer.IsIdleExportRunning = false;
			}
		}

		private static void IdleExportLoop()
		{
			int misses = 0;
			try
			{
				// PowerPoint tears the show down and repaints the editor; give
				// it that moment before the first export knocks.
				Thread.Sleep(3000);
				while (!PptAnalyzer.IsAnalyzeOver() && !PptAnalyzer._isBreakExportAsync && PptOperator.Instance != null && !PptOperator.Instance.IsShowing)
				{
					int missing = -1;
					for (int i = 1; i <= PptOperator.Instance.SlidesCount; i++)
					{
						if (!PptAnalyzer.AnalyzedIndexAndThumbnailFilePathDictionary.ContainsKey(i))
						{
							missing = i;
							break;
						}
					}
					if (missing < 0)
					{
						break;
					}
					if (PptAnalyzer.PptExportRangeImage(PptOperator.Instance.PptFilePath, null, missing, 0))
					{
						misses = 0;
						Thread.Sleep(100);
					}
					else
					{
						// A refused read is not the end of the pass - but a page
						// that never lands does not get to spin the loop forever.
						misses++;
						if (misses > 10)
						{
							Log.WriteDebug("idle export: gave up after 10 refused reads at page " + missing);
							break;
						}
						Thread.Sleep(1000);
					}
				}
				if (PptAnalyzer.IsAnalyzeOver())
				{
					Log.WriteDebug("idle export: deck complete, " + PptOperator.Instance.SlidesCount + " pages cached");
					PptAnalyzer.AddInfoToDictionary(PptOperator.Instance.PptFilePath);
				}
			}
			catch (Exception)
			{
			}
			finally
			{
				PptAnalyzer.IsIdleExportRunning = false;
			}
		}

		/// <summary>
		/// Yields the export while the user is actually interacting.
		///
		/// A slide export is a full offscreen render on PowerPoint's STA - the
		/// same thread that answers a click. Starting one while a click is in
		/// flight is the "clicked but nothing happened" race. While presses
		/// keep arriving (MouseHelper stamps them) the pass holds fire, up to a
		/// second-and-a-half long total so a page can never be starved away.
		/// Outside a show MouseHelper is stopped and the stamp goes stale, so
		/// this returns immediately.
		/// </summary>
		private static void WaitForInputGap()
		{
			for (int i = 0; i < 15; i++)
			{
				if (MouseHelper.MillisecondsSinceInput >= 1500L)
				{
					return;
				}
				Thread.Sleep(100);
			}
		}

		// Token: 0x060001FE RID: 510 RVA: 0x0000966F File Offset: 0x0000786F
		public static void StopAnalyzeAllAsyn()
		{
			if (!PptAnalyzer.IsAnalyzeAllAsyncRunning)
			{
				return;
			}
			PptAnalyzer._isBreakExportAsync = true;
		}

		// Token: 0x060001FF RID: 511 RVA: 0x0000967F File Offset: 0x0000787F
		/// <summary>
		/// Exports the thumbnails for a window of slides around one index, for a caller
		/// that only wants a neighbourhood rather than the whole deck.
		///
		/// NOT CALLED ANYWHERE. It was the entry point for the removed optional-hardware
		/// link, which asked for a page's neighbours to be pre-fetched on demand. The
		/// equivalent work now runs for the whole deck from MainWindow._pptOperatorShowBegin
		/// when the presentation starts. It is kept only so the removal is easy to
		/// reverse; delete it if that is not wanted.
		/// </summary>
		public static bool ExportRangeAround(int startIndex, int range = 4)
		{
			if (startIndex + range < 1 || startIndex - range > PptOperator.Instance.SlidesCount)
			{
				return false;
			}
			PptAnalyzer.StopAnalyzeAllAsyn();
			bool flag = PptAnalyzer.PptExportRangeImage(PptOperator.Instance.PptFilePath, null, startIndex, range);
			PptAnalyzer.CurrentShowPosition = startIndex;
			return flag;
		}

		// Token: 0x06000200 RID: 512 RVA: 0x000096B8 File Offset: 0x000078B8
		private static void AnalyzeAsync()
		{
			if (PptAnalyzer.IsAnalyzeOver())
			{
				return;
			}
			while (!PptAnalyzer.IsAnalyzeOver() && !PptAnalyzer._isBreakExportAsync && PptOperator.Instance.IsShowing)
			{
				bool flag = true;
				if (_turn)
				{
					for (int i = PptAnalyzer.CurrentShowPosition; i >= 1; i--)
					{
						if (!PptAnalyzer.AnalyzedIndexAndThumbnailFilePathDictionary.ContainsKey(i))
						{
							PptAnalyzer.WaitForInputGap();
							flag = PptAnalyzer.PptExportRangeImage(PptOperator.Instance.PptFilePath, null, i, 0);
							break;
						}
					}
				}
				else
				{
					for (int j = PptAnalyzer.CurrentShowPosition; j <= PptOperator.Instance.SlidesCount; j++)
					{
						if (!PptAnalyzer.AnalyzedIndexAndThumbnailFilePathDictionary.ContainsKey(j))
						{
							PptAnalyzer.WaitForInputGap();
							flag = PptAnalyzer.PptExportRangeImage(PptOperator.Instance.PptFilePath, null, j, 0);
							break;
						}
					}
				}
				if (!flag)
				{
					// A refused read used to end the analysis for the whole show:
					// PowerPoint can be busy at exactly the moment the first range
					// is read, and one exception left every panel empty for good.
					// Sleep and come back to it - only the while condition above,
					// the break of a stop request or the end of the show, may end
					// this loop.
					Thread.Sleep(Math.Max(PptAnalyzer._waitTime, 500));
					continue;
				}
				PptAnalyzer._count++;
				if (PptAnalyzer._count > 5)
				{
					PptAnalyzer._count = 0;
					PptAnalyzer.UpdateListBoxItem();
				}
				// A floor on the per-slide pause. The export used to throttle itself
				// by accident - every other slide ran a full GC pair - and with
				// those gone a Highest-rate run lands COM calls back to back. The
				// presentation's message queue starves, its STA starts rejecting
				// every caller, and the whole overlay becomes untouchable. The
				// pause below, together with WaitForInputGap above, keeps a
				// render gap open on PowerPoint long enough for the user's
				// clicks to land between slides.
				Thread.Sleep(Math.Max(PptAnalyzer._waitTime, 150));
				PptAnalyzer._turn = !PptAnalyzer._turn;
			}
			PptAnalyzer.UpdateListBoxItem();
			// The one collection this pipeline makes. The per-export collect
			// used to live in PptExportRangeImage's finally - a full collect-and-
			// wait TWICE PER SLIDE, which on a long deck meant the UI stopping
			// over and over while the show ran. Here the analysis is over, the
			// COM proxies it made are dead weight, and the cores belong to the
			// show again, so a single sweep is both safe and useful.
			GarbageCollection.Collect();
		}

		// Token: 0x06000201 RID: 513 RVA: 0x000097AC File Offset: 0x000079AC
		public static void InitPptAnalyzer(string pptFilePath)
		{
			List<string> list;
			PptAnalyzer.FindAlreadyAnalyzedFiles(pptFilePath, out list);
			PptAnalyzer.CreateOrGetAnalyzedIndexList();
			PptAnalyzer.CreateOrGetThumbnailFilePath();
			// Slides is a List<object> because the COM objects in it are late bound; see
			// PptOperator.Slides. The first one only supplies the slide size.
			PptAnalyzer.SetThumbnailSize(PptOperator.Instance.Slides.FirstOrDefault<object>());
		}

		// Token: 0x06000202 RID: 514 RVA: 0x000097E0 File Offset: 0x000079E0
		private static void UpdateListBoxItem()
		{
			SyntacticSugar.TryCatchAction(delegate
			{
				// Background: below WPF's input priority, so the list refresh
				// waits for the click in progress instead of front-running it.
				SyntacticSugar.CurrentDispatcherBeginInvoke(System.Windows.Threading.DispatcherPriority.Background, delegate
				{
					string[] array = new string[PptOperator.Instance.SlidesCount];
					foreach (KeyValuePair<int, string> keyValuePair in PptAnalyzer.AnalyzedIndexAndThumbnailFilePathDictionary)
					{
						int num = keyValuePair.Key - 1;
						if (num >= 0 && num <= array.Length - 1)
						{
							array[num] = keyValuePair.Value;
						}
					}
					ViewHelper.MainWindow.UpdateListBoxItem(array.ToList<string>());
				});
			}, null);
		}

		// Token: 0x06000203 RID: 515 RVA: 0x00009808 File Offset: 0x00007A08
		public static void PptExportAllImages(string pptFilePath, Action<double> exportProgress, out List<string> imageFiles)
		{
			object analyzeLocker = PptAnalyzer.AnalyzeLocker;
			lock (analyzeLocker)
			{
				PptAnalyzer._isAnalyzingAll = true;
				PptAnalyzer.IsBreakExport = false;
				imageFiles = new List<string>();
				PptAnalyzer.StopAnalyzeAllAsyn();
				try
				{
					foreach (object slide in PptOperator.Instance.Slides)
					{
						if (PptAnalyzer.IsBreakExport)
						{
							break;
						}
						int slideIndex = PptAnalyzer.ReadSlideIndex(slide);
						if (slideIndex <= 0)
						{
							continue;
						}
						if (PptAnalyzer.AnalyzedIndexAndThumbnailFilePathDictionary.ContainsKey(slideIndex))
						{
							imageFiles.Add(PptAnalyzer.AnalyzedIndexAndThumbnailFilePathDictionary[slideIndex]);
						}
						else
						{
							string path = PptAnalyzer.WriteThumbnailFile(slide, slideIndex);
							imageFiles.Add(path);
							PptAnalyzer.AnalyzedIndexAndThumbnailFilePathDictionary.TryAdd(slideIndex, path);
							if (exportProgress != null && PptOperator.Instance.SlidesCount > 0)
							{
								exportProgress((double)slideIndex / ((double)PptOperator.Instance.SlidesCount * 1.0));
							}
						}
					}
				PptAnalyzer.UpdateListBoxItem();
					if (PptAnalyzer.IsBreakExport)
					{
						imageFiles = new List<string>();
						PptAnalyzer._isAnalyzingAll = false;
						PptAnalyzer.IsBreakExport = false;
					}
					else
					{
						PptAnalyzer.AddInfoToDictionary(pptFilePath);
						imageFiles = imageFiles.OrderBy<string, int>((string x) => Convert.ToInt32(Path.GetFileNameWithoutExtension(x))).ToList<string>();
						if (exportProgress != null)
						{
							exportProgress(1.0);
						}
					}
				}
				catch (Exception)
				{
					imageFiles = new List<string>();
				}
				finally
				{
					PptAnalyzer.IsBreakExport = false;
					PptAnalyzer._isAnalyzingAll = false;
					GarbageCollection.Collect();
				}
			}
		}

		// Token: 0x06000204 RID: 516 RVA: 0x00009A40 File Offset: 0x00007C40
		public static bool PptExportRangeImage(string pptFilePath, Action<double> exportProgress, int startIndex, int range = 4)
		{
			object analyzeLocker = PptAnalyzer.AnalyzeLocker;
			bool flag2;
			lock (analyzeLocker)
			{
				int slidesCount = PptOperator.Instance.SlidesCount;
				int num = startIndex - range;
				int num2 = startIndex + range;
				if (num <= 0)
				{
					num2 -= num - 1;
				}
				if (num2 > slidesCount)
				{
					num -= num2 - slidesCount;
				}
				if (num2 < num)
				{
					flag2 = false;
				}
				else
				{
					if (num < 1)
					{
						num = 1;
					}
					if (num2 > slidesCount)
					{
						num2 = slidesCount;
					}
					try
					{
						for (int i = num - 1; i < num2; i++)
						{
							object slide = PptOperator.Instance.Slides[i];
							int slideIndex = PptAnalyzer.ReadSlideIndex(slide);
							if (slideIndex <= 0)
							{
								continue;
							}
							if (!PptAnalyzer.AnalyzedIndexAndThumbnailFilePathDictionary.ContainsKey(slideIndex))
							{
								string path = PptAnalyzer.WriteThumbnailFile(slide, slideIndex);
								PptAnalyzer.AnalyzedIndexAndThumbnailFilePathDictionary.TryAdd(slideIndex, path);
								if (exportProgress != null)
								{
									exportProgress((double)(i - num + 1) / ((double)(num2 - num + 1) * 1.0));
								}
							}
						}
						if (exportProgress != null)
						{
							exportProgress(1.0);
						}
					}
					catch (Exception ex)
					{
						// This used to swallow the failure whole - an empty slide
						// list, a busy PowerPoint, anything at all - and hand the
						// analysis a bare false with the log saying nothing. Count
						// the failures and report the first and every fortieth, so
						// a stuck export is visible without flooding the file.
						PptAnalyzer._exportFailureCount++;
						if (PptAnalyzer._exportFailureCount == 1 || PptAnalyzer._exportFailureCount % 40 == 0)
						{
							Log.WriteDebug("thumbnail export failed (" + PptAnalyzer._exportFailureCount + " total): " + ex.GetType().Name + ": " + ex.Message);
						}
						return false;
					}
					finally
					{
						// No GC here. This finally used to run the full
						// collect-and-wait twice PER EXPORTED SLIDE - on a
						// 68-page deck that is 136 global collections during a
						// live show, each stopping every thread long enough to
						// feel like a freeze, with the thumbnails themselves
						// surviving every one of them. The runtime collects on
						// its own schedule; AnalyzeAsync makes one pass when the
						// analysis is over instead.
					}
					flag2 = true;
				}
			}
			return flag2;
		}

		/// <summary>
		/// A slide's 1-based position in the presentation, or 0 if it cannot be read.
		///
		/// SlideIndex is read late bound and separately from Export because a slide
		/// whose index is unreadable must be skipped rather than written out under a
		/// made-up name that would then collide with a real one.
		/// </summary>
		internal static int ReadSlideIndex(object slide)
		{
			return ComLate.GetInt(slide, "SlideIndex");
		}

		/// <summary>
		/// Exports one slide to its thumbnail file and returns the path.
		///
		/// Export goes through IDispatch: Slide.Export is not in the hand-written
		/// interop for this project (it has no Shapes collection either, which is why
		/// the ink clearing was already late bound), and the local interface could not
		/// be cast onto a WPS slide anyway. WPS implements Export.
		///
		/// The width and height are passed explicitly, which is how the thumbnails end
		/// up at the same size as the panel regardless of the deck's own dimensions.
		/// </summary>
		internal static string WriteThumbnailFile(object slide, int slideIndex)
		{
			string path = Path.Combine(PptAnalyzer.ThumbnailFilePath, slideIndex.ToString() + PptAnalyzer.JpgSuffix);
			ComLate.Call(slide, "Export", path, PptAnalyzer.JpgSuffix, PptAnalyzer._thumbnailWidth, PptAnalyzer._thumbnailHeight);
			return path;
		}

		// Token: 0x06000205 RID: 517 RVA: 0x00009BD0 File Offset: 0x00007BD0
		private static void SetThumbnailSize(object slide)
		{
			if (slide == null)
			{
				return;
			}
			// The master supplies the deck's aspect ratio; the thumbnail width is then
			// derived from it so the panel never has to stretch anything.
			object master = ComLate.Get(slide, "Master");
			double width = ComLate.GetInt(master, "Width");
			double height = ComLate.GetInt(master, "Height");
			if (height <= 0.0)
			{
				return;
			}
			PptAnalyzer.ThumbnailRate = width / height;
			PptAnalyzer._thumbnailWidth = Convert.ToInt32((double)PptAnalyzer._thumbnailHeight * PptAnalyzer.ThumbnailRate);
		}

		// Token: 0x06000206 RID: 518 RVA: 0x00009C09 File Offset: 0x00007E09
		private static bool IsAnalyzeOver()
		{
			return PptAnalyzer.AnalyzedIndexAndThumbnailFilePathDictionary.Count == PptOperator.Instance.SlidesCount;
		}

		// Token: 0x06000207 RID: 519 RVA: 0x00009C24 File Offset: 0x00007E24
		private static void AddInfoToDictionary(string pptFilePath)
		{
			if (string.IsNullOrEmpty(PptAnalyzer._md5) || !PptAnalyzer.IsAnalyzeOver())
			{
				return;
			}
			if (!PptAnalyzer.AnalyzeFinishedFilePathDictionary.ContainsKey(PptAnalyzer._md5))
			{
				PptAnalyzer.AnalyzeFinishedFilePathDictionary.Add(PptAnalyzer._md5, pptFilePath);
			}
			if (!PptAnalyzer.AnalyzeFinishedThumbnailPathDictionary.ContainsKey(PptAnalyzer._md5))
			{
				PptAnalyzer.AnalyzeFinishedThumbnailPathDictionary.Add(PptAnalyzer._md5, PptAnalyzer.ThumbnailFilePath);
				PptAnalyzer.AppendCacheIndex(PptAnalyzer._md5, PptAnalyzer.ThumbnailFilePath);
			}
		}

		// Token: 0x06000208 RID: 520 RVA: 0x00009C8C File Offset: 0x00007E8C
		private static void CreateOrGetThumbnailFilePath()
		{
			if (string.IsNullOrEmpty(PptAnalyzer._md5))
			{
				PptAnalyzer.ThumbnailFilePath = Path.Combine(AppPath.TemPath(), Guid.NewGuid().ToString());
				if (!Directory.Exists(PptAnalyzer.ThumbnailFilePath))
				{
					Directory.CreateDirectory(PptAnalyzer.ThumbnailFilePath);
				}
				return;
			}
			if (!PptAnalyzer.ThumbnailFilePathDictionary.ContainsKey(PptAnalyzer._md5))
			{
				PptAnalyzer.ThumbnailFilePath = Path.Combine(AppPath.TemPath(), Guid.NewGuid().ToString());
				if (!Directory.Exists(PptAnalyzer.ThumbnailFilePath))
				{
					Directory.CreateDirectory(PptAnalyzer.ThumbnailFilePath);
				}
				PptAnalyzer.ThumbnailFilePathDictionary.Add(PptAnalyzer._md5, PptAnalyzer.ThumbnailFilePath);
				PptAnalyzer.AppendCacheIndex(PptAnalyzer._md5, PptAnalyzer.ThumbnailFilePath);
				return;
			}
			string text;
			PptAnalyzer.ThumbnailFilePathDictionary.TryGetValue(PptAnalyzer._md5, out text);
			PptAnalyzer.ThumbnailFilePath = text;
			PptAnalyzer.SeedAnalyzedIndexesFromFolder();
		}

		// Token: 0x06000209 RID: 521 RVA: 0x00009D58 File Offset: 0x00007F58
		private static void CreateOrGetAnalyzedIndexList()
		{
			if (string.IsNullOrEmpty(PptAnalyzer._md5))
			{
				PptAnalyzer.AnalyzedIndexAndThumbnailFilePathDictionary = new System.Collections.Concurrent.ConcurrentDictionary<int, string>();
				return;
			}
			if (!PptAnalyzer.PptMd5AndAnalyzedIndexDictionary.ContainsKey(PptAnalyzer._md5))
			{
				PptAnalyzer.AnalyzedIndexAndThumbnailFilePathDictionary = new System.Collections.Concurrent.ConcurrentDictionary<int, string>();
				PptAnalyzer.PptMd5AndAnalyzedIndexDictionary.Add(PptAnalyzer._md5, PptAnalyzer.AnalyzedIndexAndThumbnailFilePathDictionary);
				return;
			}
			System.Collections.Concurrent.ConcurrentDictionary<int, string> dictionary;
			PptAnalyzer.PptMd5AndAnalyzedIndexDictionary.TryGetValue(PptAnalyzer._md5, out dictionary);
			PptAnalyzer.AnalyzedIndexAndThumbnailFilePathDictionary = dictionary;
		}

		/// <summary>md5-to-folder lines, one per deck, so a cache survives a restart.</summary>
		private static readonly string CacheIndexPath = Path.Combine(AppPath.TemPath(), "cache-index.txt");

		private static bool _cacheIndexLoaded;

		/// <summary>
		/// Appends one md5-to-folder line to the on-disk cache index.
		///
		/// The md5 dictionaries are in-memory only, which used to mean every
		/// restart orphaned every exported folder and the next show re-exported
		/// the whole deck - the "export is slow" half of the freeze reports.
		/// A one-line-per-deck index file is all it takes for a later session
		/// to find the folder again. A locked or missing file costs nothing:
		/// the next completion re-appends.
		/// </summary>
		private static void AppendCacheIndex(string md5, string folder)
		{
			try
			{
				if (string.IsNullOrEmpty(md5) || string.IsNullOrEmpty(folder))
				{
					return;
				}
				File.AppendAllText(PptAnalyzer.CacheIndexPath, md5 + "|" + folder + "\r\n", System.Text.Encoding.UTF8);
			}
			catch (Exception)
			{
			}
		}

		/// <summary>
		/// Loads the on-disk cache index into the in-memory dictionaries, once.
		/// Entries whose folder has since been cleaned away are skipped.
		/// </summary>
		private static void LoadCacheIndex()
		{
			if (PptAnalyzer._cacheIndexLoaded)
			{
				return;
			}
			PptAnalyzer._cacheIndexLoaded = true;
			try
			{
				if (!File.Exists(PptAnalyzer.CacheIndexPath))
				{
					return;
				}
				string[] array = File.ReadAllLines(PptAnalyzer.CacheIndexPath);
				for (int i = 0; i < array.Length; i++)
				{
					string text = array[i];
					int num = text.IndexOf('|');
					if (num <= 0 || num >= text.Length - 1)
					{
						continue;
					}
					string text2 = text.Substring(0, num);
					string text3 = text.Substring(num + 1);
					if (string.IsNullOrEmpty(text2) || string.IsNullOrEmpty(text3) || !Directory.Exists(text3))
					{
						continue;
					}
					if (!PptAnalyzer.ThumbnailFilePathDictionary.ContainsKey(text2))
					{
						PptAnalyzer.ThumbnailFilePathDictionary.Add(text2, text3);
					}
					if (!PptAnalyzer.AnalyzeFinishedThumbnailPathDictionary.ContainsKey(text2))
					{
						PptAnalyzer.AnalyzeFinishedThumbnailPathDictionary.Add(text2, text3);
					}
				}
			}
			catch (Exception)
			{
			}
		}

		/// <summary>
		/// Pre-fills the analyzed-index dictionary from the jpgs already in the
		/// bound folder, so the export loop skips pages it has on disk instead
		/// of re-rendering them - which is what made a half-finished folder
		/// cost a full re-export on every run.
		/// </summary>
		private static void SeedAnalyzedIndexesFromFolder()
		{
			try
			{
				if (string.IsNullOrEmpty(PptAnalyzer.ThumbnailFilePath) || !Directory.Exists(PptAnalyzer.ThumbnailFilePath) || PptAnalyzer.AnalyzedIndexAndThumbnailFilePathDictionary == null || !PptAnalyzer.AnalyzedIndexAndThumbnailFilePathDictionary.IsEmpty)
				{
					return;
				}
				int seeded = 0;
				FileInfo[] files = new DirectoryInfo(PptAnalyzer.ThumbnailFilePath).GetFiles(PptAnalyzer.FilterName);
				for (int i = 0; i < files.Length; i++)
				{
					int num;
					if (int.TryParse(Path.GetFileNameWithoutExtension(files[i].Name), out num) && num >= 1 && PptAnalyzer.AnalyzedIndexAndThumbnailFilePathDictionary.TryAdd(num, files[i].FullName))
					{
						seeded++;
					}
				}
				if (seeded > 0)
				{
					Log.WriteDebug("cache seed: " + seeded + " thumbnails reused from disk");
				}
			}
			catch (Exception)
			{
			}
		}

		// Token: 0x0600020A RID: 522 RVA: 0x00009DC4 File Offset: 0x00007FC4
		public static bool FindAlreadyAnalyzedFiles(string filePath, out List<string> imageFiles)
		{
			bool flag;
			try
			{
				PptAnalyzer.LoadCacheIndex();
				imageFiles = new List<string>();
				if (!RegexHelper.IsLocalPath(filePath))
				{
					PptAnalyzer._md5 = null;
					flag = false;
				}
				else
				{
					if (!File.Exists(filePath) && PptAnalyzer.AnalyzeFinishedFilePathDictionary.ContainsValue(filePath))
					{
						PptAnalyzer._md5 = PptAnalyzer.AnalyzeFinishedFilePathDictionary.FirstOrDefault<KeyValuePair<string, string>>((KeyValuePair<string, string> d) => d.Value == filePath).Key;
					}
					else if (File.Exists(filePath))
					{
						PptAnalyzer._md5 = Md5SecurityUtility.GetShaFromFile(filePath);
					}
					if (string.IsNullOrEmpty(PptAnalyzer._md5))
					{
						flag = false;
					}
					else
					{
						string text;
						if (!PptAnalyzer.AnalyzeFinishedThumbnailPathDictionary.ContainsKey(PptAnalyzer._md5))
						{
							text = null;
						}
						else
						{
							text = PptAnalyzer.AnalyzeFinishedThumbnailPathDictionary.FirstOrDefault<KeyValuePair<string, string>>((KeyValuePair<string, string> d) => d.Key == PptAnalyzer._md5).Value;
						}
						string text2 = text;
						if (text2 == null)
						{
							flag = false;
						}
						else
						{
							DirectoryInfo directoryInfo = new DirectoryInfo(text2);
							if (directoryInfo.GetFiles(PptAnalyzer.FilterName).Any<FileInfo>())
							{
								imageFiles.AddRange((from x in directoryInfo.GetFiles(PptAnalyzer.FilterName)
									orderby Convert.ToInt32(Path.GetFileNameWithoutExtension(x.Name))
									select x into imageFile
									select imageFile.FullName).ToList<string>());
								PptAnalyzer.ThumbnailFilePath = text2;
								// "Any one jpg is a hit" accepted folders written half way -
								// a process killed mid-export holds 1 of 28 slides - as a
								// finished cache. The panel then built a 3-entry list for a
								// 28-page deck and the show ran against a list that did not
								// match the slides, a configuration no complete-cache run
								// ever sees (three freeze reports all show the loader
								// stopping at exactly the folder's file count). The folder
								// now counts as unfinished below its deck's length, and the
								// ordinary export path fills it; the jpgs already there are
								// kept, because the export appends to the same md5 folder.
								int expectedSlideCount = (PptOperator.Instance.ShowRangeType == PpSlideShowRangeType.ppShowSlideRange) ? PptOperator.Instance.EndingSlide : PptOperator.Instance.SlidesCount;
								if (expectedSlideCount > 0 && imageFiles.Count < expectedSlideCount)
								{
									Log.WriteDebug("partial thumbnail cache: " + imageFiles.Count + " of " + expectedSlideCount + " - treated as unfinished");
									flag = false;
								}
								else
								{
									flag = true;
								}
							}
							else
							{
								flag = false;
							}
						}
					}
				}
			}
			catch (Exception ex)
			{
				imageFiles = new List<string>();
				Log.WriteError(ex.ToString());
				flag = false;
			}
			return flag;
		}

		// Token: 0x040000F2 RID: 242
		private static bool _turn;

		// Token: 0x040000F3 RID: 243
		private static int _count;

		// Token: 0x040000F4 RID: 244
		private static string _md5;

		// Token: 0x040000F5 RID: 245
		private static Thread _thread;

		// Token: 0x040000F6 RID: 246
		public static bool IsBreakExport;

		// Token: 0x040000F7 RID: 247
		private static int _waitTime = 999;

		// Token: 0x040000F8 RID: 248
		private static bool _isAnalyzingAll;

		/// <summary>
		/// How many thumbnail exports have thrown since the process started, for
		/// the throttled report in <see cref="PptExportRangeImage"/> - the first
		/// failure and every fortieth after it get a log line, the rest only
		/// count.
		/// </summary>
		private static int _exportFailureCount;

		// Token: 0x040000F9 RID: 249
		private static bool _isBreakExportAsync;

		// Token: 0x040000FA RID: 250
		public static string ThumbnailFilePath;

		// Token: 0x040000FB RID: 251
		public static double ThumbnailRate = 1.7;

		// Token: 0x040000FC RID: 252
		private static int _thumbnailWidth = 960;

		// Token: 0x040000FD RID: 253
		private static int _thumbnailHeight = 540;

		// Token: 0x040000FE RID: 254
		public static int CurrentShowPosition = 1;

		// Token: 0x040000FF RID: 255
		public static AnalyzeRate CurrentAnalyzeRate;

		// Token: 0x04000100 RID: 256
		public static readonly string JpgSuffix = ".jpg";

		// Token: 0x04000101 RID: 257
		public static readonly string FilterName = "*.jpg";

		// Token: 0x04000102 RID: 258
		private static readonly object AnalyzeLocker = new object();

		// Token: 0x04000103 RID: 259
		//
		// Concurrent on purpose. This is written by the analysis thread (one
		// Add per exported slide, all over a live show) and read by the UI
		// thread (InitListBoxItem and UpdateListBoxItem enumerate it). A plain
		// Dictionary enumerated during a write is undefined - worst case its
		// bucket chain loops and the UI thread SPINS FOREVER, which is exactly
		// what "a 68-page deck freezes the program" was: the enumeration is on
		// the dispatcher, so nothing could interrupt it.
		public static System.Collections.Concurrent.ConcurrentDictionary<int, string> AnalyzedIndexAndThumbnailFilePathDictionary = new System.Collections.Concurrent.ConcurrentDictionary<int, string>();

		// Token: 0x04000104 RID: 260
		private static readonly Dictionary<string, string> ThumbnailFilePathDictionary = new Dictionary<string, string>();

		// Token: 0x04000105 RID: 261
		private static readonly Dictionary<string, string> AnalyzeFinishedFilePathDictionary = new Dictionary<string, string>();

		// Token: 0x04000106 RID: 262
		private static readonly Dictionary<string, string> AnalyzeFinishedThumbnailPathDictionary = new Dictionary<string, string>();

		// Token: 0x04000107 RID: 263
		private static readonly Dictionary<string, System.Collections.Concurrent.ConcurrentDictionary<int, string>> PptMd5AndAnalyzedIndexDictionary = new Dictionary<string, System.Collections.Concurrent.ConcurrentDictionary<int, string>>();
	}
}
