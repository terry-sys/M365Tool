using System.Threading.Tasks;

namespace Office365CleanupTool.Services
{
    public interface IScriptRunner
    {
        Task<ScriptExecutionResult> RunAsync(ScriptExecutionRequest request);

        void StartDetached(ScriptExecutionRequest request);
    }
}
