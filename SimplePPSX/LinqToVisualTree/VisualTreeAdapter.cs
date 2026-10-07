using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Media;

namespace LinqToVisualTree
{
	// Token: 0x02000003 RID: 3
	public class VisualTreeAdapter : ILinqTree<DependencyObject>
	{
		// Token: 0x06000007 RID: 7 RVA: 0x000020C6 File Offset: 0x000002C6
		public VisualTreeAdapter(DependencyObject item)
		{
			this._item = item;
		}

		// Token: 0x06000008 RID: 8 RVA: 0x000020D5 File Offset: 0x000002D5
		public IEnumerable<DependencyObject> Children()
		{
			int childrenCount = VisualTreeHelper.GetChildrenCount(this._item);
			int num;
			for (int i = 0; i < childrenCount; i = num + 1)
			{
				yield return VisualTreeHelper.GetChild(this._item, i);
				num = i;
			}
			yield break;
		}

		// Token: 0x17000001 RID: 1
		// (get) Token: 0x06000009 RID: 9 RVA: 0x000020E5 File Offset: 0x000002E5
		public DependencyObject Parent
		{
			get
			{
				return VisualTreeHelper.GetParent(this._item);
			}
		}

		// Token: 0x04000001 RID: 1
		private DependencyObject _item;
	}
}
