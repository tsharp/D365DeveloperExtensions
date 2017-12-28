namespace SolutionPackager.Models
{
    public class SolutionPackagerCommand
    {
        public string ToolPath { get; set; }
        public string CommandArgs { get; set; }
        public string Action { get; set; }
        public string SolutionName { get; set; }
    }
}