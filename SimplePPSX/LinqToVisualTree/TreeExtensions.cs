using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace LinqToVisualTree
{
	// Token: 0x02000005 RID: 5
	public static class TreeExtensions
	{
		// Token: 0x0600000C RID: 12 RVA: 0x000020F2 File Offset: 0x000002F2
		public static IEnumerable<DependencyObject> Descendants(this DependencyObject item)
		{
			ILinqTree<DependencyObject> linqTree = new VisualTreeAdapter(item);
			foreach (DependencyObject child in linqTree.Children())
			{
				yield return child;
				foreach (DependencyObject dependencyObject in child.Descendants())
				{
					yield return dependencyObject;
				}
			}
		}

		// Token: 0x0600000D RID: 13 RVA: 0x00002102 File Offset: 0x00000302
		public static IEnumerable<DependencyObject> DescendantsAndSelf(this DependencyObject item)
		{
			yield return item;
			foreach (DependencyObject dependencyObject in item.Descendants())
			{
				yield return dependencyObject;
			}
		}

		// Token: 0x0600000E RID: 14 RVA: 0x00002112 File Offset: 0x00000312
		public static IEnumerable<DependencyObject> Ancestors(this DependencyObject item)
		{
			ILinqTree<DependencyObject> linqTree = new VisualTreeAdapter(item);
			for (DependencyObject parent = linqTree.Parent; parent != null; parent = linqTree.Parent)
			{
				yield return parent;
				linqTree = new VisualTreeAdapter(parent);
			}
		}

		// Token: 0x0600000F RID: 15 RVA: 0x00002122 File Offset: 0x00000322
		public static IEnumerable<DependencyObject> AncestorsAndSelf(this DependencyObject item)
		{
			yield return item;
			foreach (DependencyObject dependencyObject in item.Ancestors())
			{
				yield return dependencyObject;
			}
		}

		// Token: 0x06000010 RID: 16 RVA: 0x00002132 File Offset: 0x00000332
		public static IEnumerable<DependencyObject> Elements(this DependencyObject item)
		{
			ILinqTree<DependencyObject> linqTree = new VisualTreeAdapter(item);
			foreach (DependencyObject dependencyObject in linqTree.Children())
			{
				yield return dependencyObject;
			}
		}

		// Token: 0x06000011 RID: 17 RVA: 0x00002142 File Offset: 0x00000342
		public static IEnumerable<DependencyObject> ElementsBeforeSelf(this DependencyObject item)
		{
			if (item.Ancestors().FirstOrDefault<DependencyObject>() == null)
			{
				yield break;
			}
			foreach (DependencyObject dependencyObject in item.Ancestors().First<DependencyObject>().Elements())
			{
				if (dependencyObject.Equals(item))
				{
					break;
				}
				yield return dependencyObject;
			}
		}

		// Token: 0x06000012 RID: 18 RVA: 0x00002152 File Offset: 0x00000352
		public static IEnumerable<DependencyObject> ElementsAfterSelf(this DependencyObject item)
		{
			if (item.Ancestors().FirstOrDefault<DependencyObject>() == null)
			{
				yield break;
			}
			bool afterSelf = false;
			foreach (DependencyObject child in item.Ancestors().First<DependencyObject>().Elements())
			{
				if (afterSelf)
				{
					yield return child;
				}
				if (child.Equals(item))
				{
					afterSelf = true;
				}
			}
		}

		// Token: 0x06000013 RID: 19 RVA: 0x00002162 File Offset: 0x00000362
		public static IEnumerable<DependencyObject> ElementsAndSelf(this DependencyObject item)
		{
			yield return item;
			foreach (DependencyObject dependencyObject in item.Elements())
			{
				yield return dependencyObject;
			}
		}

		// Token: 0x06000014 RID: 20 RVA: 0x00002172 File Offset: 0x00000372
		public static IEnumerable<DependencyObject> Descendants<T>(this DependencyObject item)
		{
			return (from i in item.Descendants()
				where i is T
				select i).Cast<DependencyObject>();
		}

		// Token: 0x06000015 RID: 21 RVA: 0x000021A3 File Offset: 0x000003A3
		public static IEnumerable<DependencyObject> ElementsBeforeSelf<T>(this DependencyObject item)
		{
			return (from i in item.ElementsBeforeSelf()
				where i is T
				select i).Cast<DependencyObject>();
		}

		// Token: 0x06000016 RID: 22 RVA: 0x000021D4 File Offset: 0x000003D4
		public static IEnumerable<DependencyObject> ElementsAfterSelf<T>(this DependencyObject item)
		{
			return (from i in item.ElementsAfterSelf()
				where i is T
				select i).Cast<DependencyObject>();
		}

		// Token: 0x06000017 RID: 23 RVA: 0x00002205 File Offset: 0x00000405
		public static IEnumerable<DependencyObject> DescendantsAndSelf<T>(this DependencyObject item)
		{
			return (from i in item.DescendantsAndSelf()
				where i is T
				select i).Cast<DependencyObject>();
		}

		// Token: 0x06000018 RID: 24 RVA: 0x00002236 File Offset: 0x00000436
		public static IEnumerable<DependencyObject> Ancestors<T>(this DependencyObject item)
		{
			return (from i in item.Ancestors()
				where i is T
				select i).Cast<DependencyObject>();
		}

		// Token: 0x06000019 RID: 25 RVA: 0x00002267 File Offset: 0x00000467
		public static IEnumerable<DependencyObject> AncestorsAndSelf<T>(this DependencyObject item)
		{
			return (from i in item.AncestorsAndSelf()
				where i is T
				select i).Cast<DependencyObject>();
		}

		// Token: 0x0600001A RID: 26 RVA: 0x00002298 File Offset: 0x00000498
		public static IEnumerable<DependencyObject> Elements<T>(this DependencyObject item)
		{
			return (from i in item.Elements()
				where i is T
				select i).Cast<DependencyObject>();
		}

		// Token: 0x0600001B RID: 27 RVA: 0x000022C9 File Offset: 0x000004C9
		public static IEnumerable<DependencyObject> ElementsAndSelf<T>(this DependencyObject item)
		{
			return (from i in item.ElementsAndSelf()
				where i is T
				select i).Cast<DependencyObject>();
		}
	}
}
