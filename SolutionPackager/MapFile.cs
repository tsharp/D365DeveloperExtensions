using D365DeveloperExtensions.Core.Enums;
using D365DeveloperExtensions.Core.Models;
using System.IO;
using System.Xml.Linq;

namespace SolutionPackager
{
    public class MapFile
    {
        public static void Create(SolutionPackageConfig solutionPackageConfig, string path)
        {
            // Create mapping xml with relative paths
            /*
            <?xml version="1.0" encoding="utf-8"?>
            <Mapping>
                <!-- Match specific named files to an alternate folder -->
                <!--<FileToFile map="Plugins.dll" to="..\..\Plugins\bin\**\Plugins.dll" />
                <FileToFile map="CRMDevTookitSampleWorkflow.dll" to="..\..\Workflow\bin\**\CRMDevTookitSample.Workflow.dll" />-->
                <!-- Match any file in and under WebResources to an alternate set of sub-folders -->
                <FileToPath map="PluginAssemblies\**\*.*" to="..\..\Plugins\bin\**" />
                <FileToPath map="WebResources\*.*" to="..\..\Webresources\Webresources\**" />
                <FileToPath map="WebResources\**\*.*" to="..\..\Webresources\Webresources\**" />
            </Mapping>
            */

            var mappingDoc = new XDocument();
            var mappings = new XElement("Mapping");
            mappingDoc.Add(mappings);

            foreach (var map in solutionPackageConfig.map)
            {
                if (map.map == MapTypes.file.ToString())
                    mappings.Add(new XElement("FileToFile",
                        new XAttribute("map", map.from),
                        new XAttribute("to", map.to)));

                if (map.map == MapTypes.folder.ToString())
                    mappings.Add(new XElement("Folder",
                        new XAttribute("map", map.from),
                        new XAttribute("to", map.to)));

                if (map.map == MapTypes.path.ToString())
                    mappings.Add(new XElement("FileToPath",
                        new XAttribute("map", map.from),
                        new XAttribute("to", map.to)));
            }

            string mapContent = mappingDoc.ToString();
            if (string.IsNullOrEmpty(mapContent))
                return;

            File.WriteAllText(path, mapContent);
        }
    }
}