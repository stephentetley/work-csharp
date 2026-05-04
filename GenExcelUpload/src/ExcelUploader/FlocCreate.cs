
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
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
            File.Copy(uploadTemplatePath, destination, true);
            
            using var wb = new XLWorkbook(destination);
            using var cmd = conn.CreateCommand();
            
            // "Change Request Header" tab
            string query = 
                $"""
                SELECT 
                    t."Change Request",
                    format(t."Change Request Description", {batch}, strftime(today(), '%d.%m.%y')) AS "Change Request Description",
                    t."Priority",
                    t."Due Date",
                FROM excel_uploader_floc_create.vw_change_request_header t;
                """;
            
            var ws = wb.Worksheets.Worksheet("Change Request Header");
            ws.Unprotect();
            cmd.CommandText = query;
            using var reader1 = cmd.ExecuteReader();
            var row = 6;
            while (reader1.Read())
            {
                ws.Cell(row, "A").Value = reader1.IsDBNull(0) ? "" : reader1.GetString(0);
                ws.Cell(row, "B").Value = reader1.IsDBNull(1) ? "" : reader1.GetString(1) ?? "";
                row++;
            };

            // "Change Request Notes" tab
            query = 
                $"""
                SELECT 
                    t.usmd_note
                FROM excel_uploader_floc_create.change_request_notes t;
                """;
            ws = wb.Worksheets.Worksheet("Change Request Header");
            ws.Unprotect();
            cmd.CommandText = query;
            var reader2 = cmd.ExecuteReader();
            row = 5;
            while (reader2.Read())
            {
                ws.Cell(row, "A").Value = reader2.IsDBNull(0) ? "" : reader2.GetString(0);
                row++;
            };

            // "FLOC-Functional Location" tab
            query = $"""
                SELECT 
                    t."Functional Location",
                    t."Description"
                FROM excel_uploader_floc_create.vw_functional_location t
                WHERE batch_number = {batch}
                ORDER BY t."Functional Location";
                """;
            ws = wb.Worksheets.Worksheet("FLOC-Functional Location");
            ws.Unprotect();
            cmd.CommandText = query;
            var reader3 = cmd.ExecuteReader();
            row = 6;
            while (reader3.Read())
            {
                ws.Cell(row, "A").Value = reader3.IsDBNull(0) ? "" : reader3.GetString(0);
                ws.Cell(row, "B").Value = reader3.IsDBNull(1) ? "" : reader3.GetString(1);
                row++;
            };
            wb.Save();
            Console.WriteLine($"Wrote: {destination}");
        }
    }
}
