using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImitComb.domain.Entity
{
	class StatusOperation
	{
		public string State { get; set; }
		public string NameTU { get; set; }
		public string Combination { get; set; }
		public bool IsStopAutoImitation { get; set; }
		public int CountCombinations { get; set; }
		public string EllapsedTime { get; set; }
	}
}
