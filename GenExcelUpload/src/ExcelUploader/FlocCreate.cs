
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
            string queryHeader = 
                $"""
                SELECT 
                    t."Change Request",
                    format(t."Change Request Description", {batch}, strftime(today(), '%d.%m.%y')) AS "Change Request Description",
                    t."Priority",
                    t."Due Date",
                FROM excel_uploader_floc_create.vw_change_request_header t;
                """;
            using var wb = new XLWorkbook(destination);
            var ws = wb.Worksheets.Worksheet("Change Request Header");
            ws.Unprotect();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = queryHeader;
            using var reader = cmd.ExecuteReader();
            var row = 6;
            while (reader.Read())
            {
                ws.Cell(row, "B").Value = reader.GetString(1);
                row++;
            };
            wb.Save();
            Console.WriteLine($"Wrote: {destination}");
        }
    }
}
