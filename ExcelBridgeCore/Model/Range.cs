using System.Collections.Generic;

namespace ExcelBridgeCore.Model
{
    public class Range
    {
        public string Colonne { get; set; }
        public List<CellCore> ListCellCore { get; private set; }

        public Range()
        {
            ListCellCore = new List<CellCore>();
        }

        public Range(List<CellCore> ListCellCore) : this()
        {
            this.ListCellCore = ListCellCore;
        }

        public Range(string Colonne) : this()
        {
            this.Colonne = Colonne;
        }

        public void addCellCore(CellCore cellCore) 
        {
            ListCellCore.Add(cellCore);
        }

    }
}
