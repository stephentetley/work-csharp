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
using System.Text.Json;
using ClosedXML.Excel;
using Microsoft.Extensions.FileSystemGlobbing;

using SptUtils.ReadDistBoardSchedule.Base;

namespace SptUtils.ReadDistBoardSchedule
{

    class Program
    {
        static int Main(string[] args) 
        {

            RootCommand rootCommand = new("ReadDistBoardSchedule") 
            {
                new Option<string>("--schedule_folder")
                {
                    Description = "Path to folder with Dist Board schedule xlsx files",
                    Required = true
                },              
            };
            rootCommand.SetAction(parseResult =>
            {
                string? dbScheduleFolder = parseResult.GetValue<string>("--schedule_folder");
                return ProcessWorkbooks(dbScheduleFolder);
            });
            ParseResult parseResult = rootCommand.Parse(args);
            return parseResult.Invoke();
        }

        private static int ProcessWorkbooks(string? dbScheduleFolder)
        {
            if (Directory.Exists(dbScheduleFolder)) {
                int ans = 0;
                Matcher matcher = new();
                matcher.AddInclude("*.xlsx");
                Console.WriteLine("Source: " + dbScheduleFolder);

                var matchingFiles = matcher.GetResultsInFullPath(dbScheduleFolder);
                foreach (var dbSchedule in matchingFiles)
                {
                    var res = ProcessWorkbook(dbSchedule);
                    ans += res;
                }
                return (ans > 1) ? 1 : 0;
            } 
            else
            {
                return 1;
            }

        }

        private static int ProcessWorkbook(string? dbSchedule)
        {
            if (File.Exists(dbSchedule))
            {
                var outputJsonFile = Path.ChangeExtension(dbSchedule, ".json");
                var options = new JsonWriterOptions { Indented = true };
                using var workbook = new XLWorkbook(dbSchedule);
                using var stream = File.Create(outputJsonFile);
                using var writer = new Utf8JsonWriter(stream, options);
                var fileName = Path.GetFileName(dbSchedule);
                writer.WriteStartArray();
                foreach (var worksheet in workbook.Worksheets)
                {
                    var circuits = ReadSchedule.ParseDBSchedule(fileName, worksheet);
                    if (circuits.Count > 0) Console.WriteLine("Sheet: " + worksheet.Name);
                    foreach(var circuit in circuits) circuit.WriteJson(writer);
                }
                writer.WriteEndArray();
                Console.WriteLine("Wrote: " + outputJsonFile);
                return 0;
            } else {
                Console.WriteLine("Invalid arguments");
                return 1;
            }
        }
    }
}
