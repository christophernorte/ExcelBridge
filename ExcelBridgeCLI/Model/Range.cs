using System.Collections.Generic;

namespace ExcelBridgeCli.Model
{
    class Range
    {
        public List<Cell> Values { get; private set; }

        public Range(List<Cell> Values)
        {
            this.Values = Values;
        }
    }
}
