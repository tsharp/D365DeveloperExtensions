namespace PluginTraceViewer.ViewModels
{
    public interface IFilterProperty
    {
        string Name { get; set; }
        string Value { get; set; }
        bool IsSelected { get; set; }
    }
}