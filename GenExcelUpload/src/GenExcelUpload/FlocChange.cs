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



using ClosedXML.Excel;
using Microsoft.Data.Sqlite;

namespace SptUtils.GenExcelUpload
{
    public static class FlocChange
    {
        public static void WriteFlocChangeUpload(SqliteConnection conn, AppSettings appSettings, string outputFolder, string? nameRoot)
        {
            string MakeOutputName(int i){
                if (nameRoot is string root)
                {
                    var name1 = $"{root}_FLOC_CHANGE_{i:D2}.xlsx";
                    return Path.Combine(outputFolder, name1);
                }
                else
                {
                    var name1 =  $"FLOC_CHANGE_{i:D2}.xlsx";
                    return Path.Combine(outputFolder, name1);
                }
            };
            using var cmd = conn.CreateCommand();
            var query = 
                """
                with cte as (
                    select max(batch_number) as max_batch from floc_change_functional_location
                    union
                    select max(batch_number) as max_batch from floc_change_classification
                )
                select max(max_batch) from cte;
                """;
            cmd.CommandText = query;
            if (cmd.ExecuteScalar() is long maxBatch)
            {
                for (int i = 1; i <= maxBatch; i++)
                {
                    Console.WriteLine($"Index: {i}");
                    var uploadTemplatePath = appSettings.FlocChangeTemplatePath;

                    var dest = MakeOutputName(i);
                    GenExcelUpload1(conn, uploadTemplatePath, dest, i);
                }
            }
        }
        private static void GenExcelUpload1(SqliteConnection conn, string uploadTemplatePath, string dest, int batch)
        {
            File.Copy(uploadTemplatePath, dest, true);
            
            using var wb = new XLWorkbook(dest);
            
            
            // "Change Request Header" tab
            string query = 
                $"""
                SELECT 
                    t.change_request,
                    t.change_request_description,
                    t.priority,
                    t.due_date
                FROM floc_change_change_request_header t;
                """;
            var ws = wb.Worksheets.Worksheet("Change Request Header");
            ws.Unprotect();
            using var cmd1 = new SqliteCommand(query, conn);
            using var reader1 = cmd1.ExecuteReader();
            var row = 6;
            while (reader1.Read())
            {
                if (!reader1.IsDBNull(0)) ws.Cell(row, "A").Value = reader1.GetString(0);   // Change Request
                if (!reader1.IsDBNull(1)) ws.Cell(row, "B").Value = reader1.GetString(1);   // Change Request Description
                if (!reader1.IsDBNull(2)) ws.Cell(row, "C").Value = reader1.GetString(2);   // Priority
                if (!reader1.IsDBNull(3)) ws.Cell(row, "D").Value = reader1.GetString(3);   // Due Date
                row++;
            };

            // "Change Request Notes" tab
            query = 
                $"""
                SELECT 
                    t.notes
                FROM floc_change_change_request_notes t;
                """;
            ws = wb.Worksheets.Worksheet("Change Request Header");
            ws.Unprotect();
            using var cmd2 = new SqliteCommand(query, conn);
            var reader2 = cmd2.ExecuteReader();
            row = 5;
            while (reader2.Read())
            {
                if (!reader2.IsDBNull(0)) ws.Cell(row, "A").Value = reader2.GetString(0);   // usmd_note
                row++;
            };

            // "FLOC-Functional Location" tab
            query = $"""
                SELECT 
                    t.functional_location,
                    t.floc_description,
                    t.inactive,
                    t.object_type,
                    t.authoriz_group,
                    t.gross_weight,
                    t.unit_of_weight,
                    inventory_no,
                    t.size_dimens,
                    t.start_up_date,
                    t.acquisition_value,
                    t.currency,
                    t.acquistion_date,
                    t.manufacturer,
                    t.model_number,
                    t.manuf_part_no,
                    t.manuf_serial_no,
                    t.manuf_country,
                    t.construct_year,
                    t.construct_mth,
                    t.maint_plant,
                    t.location, 
                    t.room,
                    t.plant_section,
                    t.work_center,
                    t.abc_indic,
                    t.sort_field,
                    t.business_area,
                    t.asset,
                    t.sub_number,
                    t.cost_center,
                    t.wbs_element,
                    t.standg_order,
                    t.settlement_order,
                    t.planning_plant,
                    t.planner_group,
                    t.main_work_ctr,
                    t.plnt_work_center,
                    t.catalog_profile,
                    t.sup_funct_loc,
                    t.position,
                    t.installation_allowed,
                    t.single_inst,
                    t.construction_type,
                    t.status_profile,
                    t.user_status,
                    t.status_of_an_object,
                    t.status_without_stsno,
                    t.begin_guarantee_c,
                    t.warranty_end_c,
                    t.master_warranty_c,
                    t.inherit_warranty_c,
                    t.pass_on_warranty_c,
                    t.begin_guarantee_v,
                    t.warranty_end_v,
                    t.master_warranty_v,
                    t.inherit_warranty_v,
                    t.pass_on_warranty_v,
                    t.sales_org,
                    t.distr_channel,
                    t.division,
                    t.sales_office,
                    t.sales_group
                FROM floc_change_functional_location t
                WHERE batch_number = {batch}
                ORDER BY t.functional_location;
                """;
            ws = wb.Worksheets.Worksheet("FLOC-Functional Location");
            ws.Unprotect();
            using var cmd3 = new SqliteCommand(query, conn);
            var reader3 = cmd3.ExecuteReader();
            row = 6;
            while (reader3.Read())
            {
                if (!reader3.IsDBNull(0)) ws.Cell(row, "A").Value = reader3.GetString(0);   // Functional Location
                if (!reader3.IsDBNull(1)) ws.Cell(row, "B").Value = reader3.GetString(1);   // Description
                if (!reader3.IsDBNull(2)) ws.Cell(row, "C").Value = reader3.GetString(2);   // FunctLocCat
                if (!reader3.IsDBNull(3)) ws.Cell(row, "D").Value = reader3.GetString(3);   // StrIndicator
                if (!reader3.IsDBNull(4)) ws.Cell(row, "E").Value = reader3.GetString(4);   // Inactive
                if (!reader3.IsDBNull(5)) ws.Cell(row, "F").Value = reader3.GetString(5);   // Object type
                if (!reader3.IsDBNull(6)) ws.Cell(row, "G").Value = reader3.GetString(6);   // AuthorizGroup
                if (!reader3.IsDBNull(7)) ws.Cell(row, "H").Value = reader3.GetDecimal(7);   // Gross Weight
                if (!reader3.IsDBNull(8)) ws.Cell(row, "I").Value = reader3.GetString(8);   // Unit of weight
                if (!reader3.IsDBNull(9)) ws.Cell(row, "J").Value = reader3.GetString(9);   // Inventory no
                if (!reader3.IsDBNull(10)) ws.Cell(row, "K").Value = reader3.GetString(10);   // Size/dimens
                if (!reader3.IsDBNull(11)) ws.Cell(row, "L").Value = reader3.GetString(11);   // Start-up date
                if (!reader3.IsDBNull(12)) ws.Cell(row, "M").Value = reader3.GetDecimal(12);   // AcquisitionValue
                if (!reader3.IsDBNull(13)) ws.Cell(row, "N").Value = reader3.GetString(13);   // Currency
                if (!reader3.IsDBNull(14)) ws.Cell(row, "O").Value = reader3.GetString(14);   // Acquistion date
                if (!reader3.IsDBNull(15)) ws.Cell(row, "P").Value = reader3.GetString(15);   // Manufacturer
                if (!reader3.IsDBNull(16)) ws.Cell(row, "Q").Value = reader3.GetString(16);   // Model number
                if (!reader3.IsDBNull(17)) ws.Cell(row, "R").Value = reader3.GetString(17);   // ManufPartNo
                if (!reader3.IsDBNull(18)) ws.Cell(row, "S").Value = reader3.GetString(18);   // ManufSerialNo
                if (!reader3.IsDBNull(19)) ws.Cell(row, "T").Value = reader3.GetString(19);   // ManufCountry
                if (!reader3.IsDBNull(20)) ws.Cell(row, "U").Value = reader3.GetString(20);   // ConstructYear
                if (!reader3.IsDBNull(21)) ws.Cell(row, "V").Value = reader3.GetString(21);   // ConstructMth
                if (!reader3.IsDBNull(22)) ws.Cell(row, "W").Value = reader3.GetString(22);   // MaintPlant
                if (!reader3.IsDBNull(23)) ws.Cell(row, "X").Value = reader3.GetString(26);   // Location
                if (!reader3.IsDBNull(24)) ws.Cell(row, "Y").Value = reader3.GetString(24);   // Room
                if (!reader3.IsDBNull(25)) ws.Cell(row, "Z").Value = reader3.GetString(25);   // Plant section
                if (!reader3.IsDBNull(26)) ws.Cell(row, "AA").Value = reader3.GetString(26);   // Work center
                if (!reader3.IsDBNull(27)) ws.Cell(row, "AB").Value = reader3.GetString(27);   // ABC indic
                if (!reader3.IsDBNull(28)) ws.Cell(row, "AC").Value = reader3.GetString(28);   // Sort field
                if (!reader3.IsDBNull(29)) ws.Cell(row, "AD").Value = reader3.GetString(29);   // Business Area
                if (!reader3.IsDBNull(30)) ws.Cell(row, "AE").Value = reader3.GetString(30);   // Asset
                if (!reader3.IsDBNull(31)) ws.Cell(row, "AF").Value = reader3.GetString(31);   // Sub-number
                if (!reader3.IsDBNull(32)) ws.Cell(row, "AG").Value = reader3.GetString(32);   // Cost Center
                if (!reader3.IsDBNull(33)) ws.Cell(row, "AH").Value = reader3.GetString(33);   // WBS Element
                if (!reader3.IsDBNull(34)) ws.Cell(row, "AI").Value = reader3.GetString(34);   // StandgOrder
                if (!reader3.IsDBNull(35)) ws.Cell(row, "AJ").Value = reader3.GetString(35);   // SettlementOrder
                if (!reader3.IsDBNull(36)) ws.Cell(row, "AK").Value = reader3.GetString(36);   // Planning plant
                if (!reader3.IsDBNull(37)) ws.Cell(row, "AL").Value = reader3.GetString(37);   // Planner group
                if (!reader3.IsDBNull(38)) ws.Cell(row, "AM").Value = reader3.GetString(38);   // Main WorkCtr
                if (!reader3.IsDBNull(39)) ws.Cell(row, "AN").Value = reader3.GetString(39);   // Plnt WorkCenter
                if (!reader3.IsDBNull(40)) ws.Cell(row, "AO").Value = reader3.GetString(40);   // Catalog profile
                if (!reader3.IsDBNull(41)) ws.Cell(row, "AP").Value = reader3.GetString(41);   // SupFunctLoc
                if (!reader3.IsDBNull(42)) ws.Cell(row, "AQ").Value = reader3.GetString(42);   // Position
                if (!reader3.IsDBNull(43)) ws.Cell(row, "AR").Value = reader3.GetString(43);   // Ref. Location
                if (!reader3.IsDBNull(44)) ws.Cell(row, "AS").Value = reader3.GetString(44);   // Installation Allowed
                if (!reader3.IsDBNull(45)) ws.Cell(row, "AT").Value = reader3.GetString(45);   // Single Inst.
                if (!reader3.IsDBNull(46)) ws.Cell(row, "AU").Value = reader3.GetString(46);   // Construction type
                if (!reader3.IsDBNull(47)) ws.Cell(row, "AV").Value = reader3.GetString(47);   // Status Profile
                if (!reader3.IsDBNull(48)) ws.Cell(row, "AW").Value = reader3.GetString(48);   // User Status
                if (!reader3.IsDBNull(49)) ws.Cell(row, "AX").Value = reader3.GetString(49);   // Status of an object
                if (!reader3.IsDBNull(50)) ws.Cell(row, "AY").Value = reader3.GetString(50);   // Status without stsno
                if (!reader3.IsDBNull(51)) ws.Cell(row, "AZ").Value = reader3.GetString(51);   // Begin guarantee(C)
                if (!reader3.IsDBNull(52)) ws.Cell(row, "BA").Value = reader3.GetString(52);   // Warranty end(C)
                if (!reader3.IsDBNull(53)) ws.Cell(row, "BB").Value = reader3.GetString(53);   // Master Warranty(C)
                if (!reader3.IsDBNull(54)) ws.Cell(row, "BC").Value = reader3.GetString(54);   // InheritWarranty(C)
                if (!reader3.IsDBNull(55)) ws.Cell(row, "BD").Value = reader3.GetString(55);   // Pass on warranty(C)
                if (!reader3.IsDBNull(56)) ws.Cell(row, "BE").Value = reader3.GetString(56);   // Begin guarantee(V)
                if (!reader3.IsDBNull(57)) ws.Cell(row, "BF").Value = reader3.GetString(57);   // Warranty end(V)
                if (!reader3.IsDBNull(58)) ws.Cell(row, "BG").Value = reader3.GetString(58);   // Master Warranty(V)
                if (!reader3.IsDBNull(59)) ws.Cell(row, "BH").Value = reader3.GetString(59);   // InheritWarranty(V)
                if (!reader3.IsDBNull(60)) ws.Cell(row, "BI").Value = reader3.GetString(60);   // Pass on warranty(V)
                if (!reader3.IsDBNull(61)) ws.Cell(row, "BJ").Value = reader3.GetString(61);   // Sales Org
                if (!reader3.IsDBNull(62)) ws.Cell(row, "BK").Value = reader3.GetString(62);   // Distr. Channel
                row++;
            };

            // 'FLOC-Classification' tab
            query = 
                $"""
                SELECT 
                    t.functional_location,
                    t.class,
                    t.characteristics,
                    t.char_value,
                    t.class_delete_ind
                FROM floc_change_classification t
                WHERE t.batch_number = {batch};
                """;
            ws = wb.Worksheets.Worksheet("FLOC-Classification");
            ws.Unprotect();
            using var cmd4 = new SqliteCommand(query, conn);
            var reader4 = cmd4.ExecuteReader();
            row = 5;
            while (reader4.Read())
            {
                if (!reader4.IsDBNull(0)) ws.Cell(row, "A").Value = reader4.GetString(0);   // Functional Location
                if (!reader4.IsDBNull(1)) ws.Cell(row, "B").Value = reader4.GetString(1);   // Class
                if (!reader4.IsDBNull(2)) ws.Cell(row, "C").Value = reader4.GetString(2);   // Characteristics
                if (!reader4.IsDBNull(3)) ws.Cell(row, "D").Value = reader4.GetString(3);   // Char Value
                if (!reader4.IsDBNull(4)) ws.Cell(row, "E").Value = reader4.GetString(4);   // Class Delete Ind
                row++;
            };
            wb.Save();
        }

    }
}
