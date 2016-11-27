using CommandLine;
using CommandLine.Text;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelBridgeCli.Argument
{
    public class Options
    {
      
        [Option('m', "mode", Required = true, HelpText = "Values are cell, range, file")]
        public string Mode { get; set; }
        
        [Option('e', "excelFilePath", Required = true, HelpText = "Path to the excel file to edit")]
        public string ExcelFilePath { get; set; }
       
        [Option('s', "sheetName", HelpText = "Sheet aimed to be edited")]
        public string SheetName { get; set; }
        
        [Option('c', "cell", HelpText = "Cell coordinated, ex : C3")]
        public string Cell { get; set; }

        [Option('v', "value", HelpText = "Value to past in the selected cell")]
        public string Value { get; set; }

        [ParserState]
        public IParserState LastParserState { get; set; }

        [HelpOption]
        public string GetUsage()
        {
            return HelpText.AutoBuild(this,(HelpText current) => HelpText.DefaultParsingErrorsHandler(this, current));
        }
    }

}
