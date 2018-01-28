namespace D365DeveloperExtensions.Core.DataGrid
{
    public interface IFilterProperty
    {
        string Name { get; set; }
        string Value { get; set; }
        bool IsSelected { get; set; }
    }
}