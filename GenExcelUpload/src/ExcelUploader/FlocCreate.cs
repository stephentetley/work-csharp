

namespace SptUtils.GenExcelUpload {


    public class FlocCreate
    {
        string conn;

        public FlocCreate(string con)
        {
            conn = con;    
        }
        
        public string InsertTitleFormatString(string titleFormat) {
            
            string s1 = "DELETE FROM excel_uploader_floc_create.change_request_header";
            string updateStmt = 
                $"""
                INSERT INTO excel_uploader_floc_create.change_request_header BY NAME
                SELECT '{titleFormat}' AS change_request_decription;
                """;
            return s1 + Environment.NewLine + updateStmt;
        }
    }

}