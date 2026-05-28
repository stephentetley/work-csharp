// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at

// http://www.apache.org/licenses/LICENSE-2.0

// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.

using DuckDB.NET.Data;

namespace SptUtils.FlocDelta.Base 
{
    public static class RunFlocDelta
    {
        
        private static int ExecScript(DuckDBConnection conn, string path)
        {
            using var command = conn.CreateCommand();
            var statements = File.ReadAllText(path);
            command.CommandText = statements;
            var ans = command.ExecuteNonQuery();
            return ans;
        }
        
        // returns number of rows afffected (includes multiple non-queries)
        public static int GenerateSQLiteDB(AppSettings appSettings, string worklistXlsx, string sqliteOutput)
        {
            
            using var conn = new DuckDBConnection("Data Source=:memory:");
            conn.Open();
            using var command = conn.CreateCommand();
            
            command.CommandText = $"INSTALL sqlite; LOAD sqlite;";
            var i = command.ExecuteNonQuery();

            command.CommandText = $"INSTALL ducklake; LOAD ducklake;";
            i = command.ExecuteNonQuery();

            command.CommandText = $"attach '{appSettings.AssetLakeConnectionString}' as asset_lake (READONLY);";
            i += command.ExecuteNonQuery();
            
            command.CommandText = $"attach '{sqliteOutput}' as sqlite_db (type sqlite);";
            i += command.ExecuteNonQuery();

            var sqlPath = Path.Combine(appSettings.WorkSQLRoot, "Scripts2/excel_uploader/create_sqlite_tables.sql");
            i += ExecScript(conn, sqlPath);

            

            command.CommandText = $"detach sqlite_db;";
            i += command.ExecuteNonQuery();

            command.CommandText = $"detach asset_lake;";
            i += command.ExecuteNonQuery();

            return i;
        }
    }
}

