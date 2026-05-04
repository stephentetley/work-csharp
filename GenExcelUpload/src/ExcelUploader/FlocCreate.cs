
using System.Security.Cryptography;
using DuckDB.NET.Data;

namespace SptUtils.GenExcelUpload {


    public class FlocCreate
    {
        DuckDBConnection conn;

        public FlocCreate(DuckDBConnection con)
        {
            conn = con;    
        }
        
        public int InsertTitleFormatString(string titleFormat) {
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "DELETE FROM excel_uploader_floc_create.change_request_header";
            var ans = cmd.ExecuteNonQuery();

            string updateStmt = 
                $"""
                INSERT INTO excel_uploader_floc_create.change_request_header BY NAME
                SELECT '{titleFormat}' AS change_request_decription;
                """;
            cmd.CommandText = updateStmt;
            ans = cmd.ExecuteNonQuery();
            return ans;
        }


        public void WriteFlocCreateUpload(string uploadTemplatePath, string dest)
        {
                // rs = con.execute("SELECT max(batch_number) FROM excel_uploader_floc_create.batch_worklist;").fetchone();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT max(batch_number) FROM excel_uploader_floc_create.batch_worklist;";
            int maxBatchNum = Convert.ToInt32(cmd.ExecuteScalar());
            for (int i = 1; i <= maxBatchNum; i++)
            {
                GenExcelUpload1(uploadTemplatePath, dest, i);
            }

        }

        private void GenExcelUpload1(string uploadTemplatePath, string dest, int batch)
        {
            var destination = string.Format(dest, batch.ToString("00"));
            Console.WriteLine($"Output file: {destination}");
        }
    }
}
