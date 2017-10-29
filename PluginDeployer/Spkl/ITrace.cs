namespace PluginDeployer.Spkl
{
    public interface ITrace
    {
        void WriteLine(string format, params object[] args);
        void Write(string format, params object[] args);
    }
}
