using System;
using System.Collections.Generic;
using System.Windows;

namespace LinqToVisualTree
{
	// Token: 0x02000006 RID: 6
	public static class EnumerableTreeExtensions
	{
		// Token: 0x0600001C RID: 28 RVA: 0x000022FA File Offset: 0x000004FA
		private static IEnumerable<DependencyObject> DrillDown(this IEnumerable<DependencyObject> items, Func<DependencyObject, IEnumerable<DependencyObject>> function)
		{
			foreach (DependencyObject dependencyObject in items)
			{
				foreach (DependencyObject dependencyObject2 in function(dependencyObject))
				{
					yield return dependencyObject2;
				}
			}
		}

		// Token: 0x0600001D RID: 29 RVA: 0x00002311 File Offset: 0x00000511
		public static IEnumerable<DependencyObject> DrillDown<T>(this IEnumerable<DependencyObject> items, Func<DependencyObject, IEnumerable<DependencyObject>> function) where T : DependencyObject
		{
			foreach (DependencyObject dependencyObject in items)
			{
				foreach (DependencyObject dependencyObject2 in function(dependencyObject))
				{
					if (dependencyObject2 is T)
					{
						yield return (T)((object)dependencyObject2);
					}
				}
			}
		}

		// Token: 0x0600001E RID: 30 RVA: 0x00002328 File Offset: 0x00000528
		public static IEnumerable<DependencyObject> Descendants(this IEnumerable<DependencyObject> items)
		{
			return items.DrillDown((DependencyObject i) => i.Descendants());
		}

		// Token: 0x0600001F RID: 31 RVA: 0x0000234F File Offset: 0x0000054F
		public static IEnumerable<DependencyObject> DescendantsAndSelf(this IEnumerable<DependencyObject> items)
		{
			return items.DrillDown((DependencyObject i) => i.DescendantsAndSelf());
		}

		// Token: 0x06000020 RID: 32 RVA: 0x00002376 File Offset: 0x00000576
		public static IEnumerable<DependencyObject> Ancestors(this IEnumerable<DependencyObject> items)
		{
			return items.DrillDown((DependencyObject i) => i.Ancestors());
		}

		// Token: 0x06000021 RID: 33 RVA: 0x0000239D File Offset: 0x0000059D
		public static IEnumerable<DependencyObject> AncestorsAndSelf(this IEnumerable<DependencyObject> items)
		{
			return items.DrillDown((DependencyObject i) => i.AncestorsAndSelf());
		}

		// Token: 0x06000022 RID: 34 RVA: 0x000023C4 File Offset: 0x000005C4
		public static IEnumerable<DependencyObject> Elements(this IEnumerable<DependencyObject> items)
		{
			return items.DrillDown((DependencyObject i) => i.Elements());
		}

		// Token: 0x06000023 RID: 35 RVA: 0x000023EB File Offset: 0x000005EB
		public static IEnumerable<DependencyObject> ElementsAndSelf(this IEnumerable<DependencyObject> items)
		{
			return items.DrillDown((DependencyObject i) => i.ElementsAndSelf());
		}

		// Token: 0x06000024 RID: 36 RVA: 0x00002412 File Offset: 0x00000612
		public static IEnumerable<DependencyObject> Descendants<T>(this IEnumerable<DependencyObject> items) where T : DependencyObject
		{
			return items.DrillDown<T>((DependencyObject i) => i.Descendants());
		}

		// Token: 0x06000025 RID: 37 RVA: 0x00002439 File Offset: 0x00000639
		public static IEnumerable<DependencyObject> DescendantsAndSelf<T>(this IEnumerable<DependencyObject> items) where T : DependencyObject
		{
			return items.DrillDown<T>((DependencyObject i) => i.DescendantsAndSelf());
		}

		// Token: 0x06000026 RID: 38 RVA: 0x00002460 File Offset: 0x00000660
		public static IEnumerable<DependencyObject> Ancestors<T>(this IEnumerable<DependencyObject> items) where T : DependencyObject
		{
			return items.DrillDown<T>((DependencyObject i) => i.Ancestors());
		}

		// Token: 0x06000027 RID: 39 RVA: 0x00002487 File Offset: 0x00000687
		public static IEnumerable<DependencyObject> AncestorsAndSelf<T>(this IEnumerable<DependencyObject> items) where T : DependencyObject
		{
			return items.DrillDown<T>((DependencyObject i) => i.AncestorsAndSelf());
		}

		// Token: 0x06000028 RID: 40 RVA: 0x000024AE File Offset: 0x000006AE
		public static IEnumerable<DependencyObject> Elements<T>(this IEnumerable<DependencyObject> items) where T : DependencyObject
		{
			return items.DrillDown<T>((DependencyObject i) => i.Elements());
		}

		// Token: 0x06000029 RID: 41 RVA: 0x000024D5 File Offset: 0x000006D5
		public static IEnumerable<DependencyObject> ElementsAndSelf<T>(this IEnumerable<DependencyObject> items) where T : DependencyObject
		{
			return items.DrillDown<T>((DependencyObject i) => i.ElementsAndSelf());
		}
	}
}
