using System;
using System.Collections.Generic;

namespace LinqToVisualTree
{
	// Token: 0x02000004 RID: 4
	public interface ILinqTree<T>
	{
		// Token: 0x0600000A RID: 10
		IEnumerable<T> Children();

		// Token: 0x17000002 RID: 2
		// (get) Token: 0x0600000B RID: 11
		T Parent { get; }
	}
}
