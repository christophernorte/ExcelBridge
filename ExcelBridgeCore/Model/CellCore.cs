namespace ExcelBridgeCore.Model
{
    public class CellCore
    {
        public uint RawIndex { get; private set; }
        public string Value { get; private set; }

        public CellCore(uint RawIndex, string Value)
        {
            this.RawIndex = RawIndex;
            this.Value = Value;
        }
    }
}
