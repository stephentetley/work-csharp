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
using Microsoft.Extensions.Configuration;
using Microsoft.Data.Sqlite;


namespace SptUtils.GenExcelUpload
{

    class Program
    {

        static int Main(string[] args) 
        {

            var builder = new ConfigurationBuilder()
                .AddJsonFile("appsettings.json")
                .Build();
            
            var appSettings = builder.GetSection("AppSettings").Get<AppSettings>();

            RootCommand rootCommand = new("GenFileUpload") 
            {
                new Option<string>("--database_file")
                {
                    Description = "The SQLite database file",
                    Required = true
                },
                new Option<string>("--output_folder")
                {
                    Description = "Output foldere for *.xlsx files",
                    Required = true
                },
                new Option<string>("--output_basename")
                {
                    Description = "Base file name for output - the name will be suffixed with the upload format and batch number",
                    Required = false
                },
            };
            rootCommand.SetAction(parseResult =>
            {

                if (appSettings != null 
                        && parseResult.GetValue<string>("--database_file") is string databaseFile 
                        && parseResult.GetValue<string>("--output_folder") is string outputFolder)
                {
                    Console.WriteLine($"GenExcelUpload:");
                    Console.WriteLine($"{databaseFile}");
                    Console.WriteLine($"{outputFolder}");
                    
                    // TODO validate databaseFile
                    string connstr = $"DataSource = {databaseFile};ACCESS_MODE=READ_WRITE;";
                    using var connection = new SqliteConnection($"Data Source={databaseFile}");
                    connection.Open();

                    var nameRoot = parseResult.GetValue<string>("--output_basename");
                    
                    FlocCreate.WriteFlocCreateUpload(connection, appSettings, outputFolder, nameRoot);
                        
                    // var equiMake = new EquiCreate(connection);
                    EquiCreate.WriteEquiCreateUpload(connection, appSettings, outputFolder, nameRoot);
                    return 0;
                }
                else
                {
                    return 1;
                }
                
            });

            ParseResult parseResult = rootCommand.Parse(args);
            return parseResult.Invoke();
           
        }
    }

}
