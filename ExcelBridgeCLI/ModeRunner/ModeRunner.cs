using ExcelBridgeCli.Argument;

namespace ExcelBridgeCli.ModeRunner
{
    public interface IModeRunner
    {
        ModeRunnerResponse Run();

        void ValidateOptions();
    }
}
