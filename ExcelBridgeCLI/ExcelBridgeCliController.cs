using ExcelBridgeCli.Argument;
using ExcelBridgeCli.ModeRunner;

namespace ExcelBridgeCli
{
    public class ExcelBridgeCliController
    {
        public void Start(string[] args)
        {
            var options = new Options();
            if (CommandLine.Parser.Default.ParseArguments(args, options))
            {
                IModeRunner modeRunner = ModeRunnerFactory.getRunner(options.Mode, options);

                modeRunner.ValidateOptions();
                modeRunner.Run();
            }
        }
    }
}
