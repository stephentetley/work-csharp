// Copyright 2026 Stephen Tetley

// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at

// http://www.apache.org/licenses/LICENSE-2.0

// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.


using System.CommandLine;
using ClosedXML.Excel;

namespace SptUtils.ReadAsFitted {

    class Program
    {
        static int Main(string[] args) 
        {

            RootCommand rootCommand = new("ReadAsFitted") 
            {
                new Option<string>("--as_fitted_file")
                {
                    Description = "Path to Ass Fitted xlsx file",
                    Required = true
                },
            };
            rootCommand.SetAction(parseResult =>
            {
                string? asFitted = parseResult.GetValue<string>("--as_fitted_file");
                if (File.Exists(asFitted))
                {
                    using var workbook = new XLWorkbook(asFitted);
                    foreach(var worksheet in workbook.Worksheets)
                    {
                        var reader = new ReadTestSheet(worksheet);

                        Console.WriteLine("Sheet: " + worksheet.Name + " " + reader.isTestSheet());
                        if (reader.isTestSheet())
                        {
                            reader.readTest();
                        }
                    } 

                    Console.WriteLine("Source: " + asFitted);
                    return 0;
                } else {
                    Console.WriteLine("Invalid arguments");
                    return 1;
                }
            });
            ParseResult parseResult = rootCommand.Parse(args);
            return parseResult.Invoke();
        }
    }
}

