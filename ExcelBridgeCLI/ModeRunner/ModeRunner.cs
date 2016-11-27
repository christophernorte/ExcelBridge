using ExcelBridgeCli.Argument;

namespace ExcelBridgeCli.ModeRunner
{
    public interface ModeRunner
    {
        ModeRunnerResponse Run(Options option);
    }
}
