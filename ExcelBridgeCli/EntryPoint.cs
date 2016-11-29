using ExcelBridgeCli.Exceptions;

namespace ExcelBridgeCli
{
    class EntryPoint
    {
        static void Main(string[] args)
        {
            try {
                ExcelBridgeCliController ExcelBridgeController = new ExcelBridgeCliController();
                ExcelBridgeController.Start(args);
            }
            catch (CliArgumentMissing e)
            {
                System.Console.Error.WriteLine(e.Message);
            }catch(CliArgumentCellBadFormat e)
            {
                System.Console.Error.WriteLine(e.Message);
            }
        }

    }
}