using System;
using CommandLine;

namespace WordCompressPNG
{
    public class Options
    {
        [Option('i', "input", Required = true, HelpText = "Input Word DOCX file (complete path and extension)")]
        public string? InputPath { get; set; }

        [Option('o', "output", Required = false, HelpText = "Output Word DOCX file (complete path and extension)")]
        public string? OutputPath { get; set; }

        [Option('w', "overwrite", Required = false, HelpText = "Overwrite existing Input Word DOCX file")]
        public bool Overwrite { get; set; }


        [Option('k', "apikey", Required = false, HelpText = "CloudConvert API Key")]
        public string? ApiKey { get; set; }

        [Option('s', "sandbox", Required = false, HelpText = "Enable sandbox mode (development keys)")]
        public bool Sandbox { get; set; }

        [Option("maxfiles", Required =false,HelpText ="Maximum number of PNG files to compress (0=all)")]
        public int MaxFiles { get; set; }
    }
}
