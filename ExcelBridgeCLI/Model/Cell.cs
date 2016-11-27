namespace ExcelBridgeCli.Model
{
    class Cell
    {
        public string Position { get; private set; }
        public string Value { get; private set; }

        public Cell(string Position, string Value)
        {
            this.Position = Position;
            this.Value = Value;
        }
    }
}
