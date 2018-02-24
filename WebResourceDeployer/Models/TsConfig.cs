namespace WebResourceDeployer.Models
{
    public class TsConfig
    {
        public bool compileOnSave { get; set; }
        public Compileroptions compilerOptions { get; set; }
        public string[] include { get; set; }
    }

    public class Compileroptions
    {
        public string outDir { get; set; }
        public string target { get; set; }
        public bool sourceMap { get; set; }
    }
}