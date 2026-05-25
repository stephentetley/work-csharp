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

namespace SptUtils.GenExcelUpload {


    public static class EquiChange
    {

        public static void WriteEquiChangeUpload(SqliteConnection conn, AppSettings appSettings, string outputFolder, string? nameRoot)
        {
            string MakeOutputName(int i){
                if (nameRoot is string root)
                {
                    var name1 = $"{root}_EQUI_CHANGE_{i:D2}.xlsx";
                    return Path.Combine(outputFolder, name1);
                }
                else
                {
                    var name1 =  $"EQUI_CHANGE_{i:D2}.xlsx";
                    return Path.Combine(outputFolder, name1);
                }
            };
            using var cmd = conn.CreateCommand();
            var query = 
                """
                with cte as (
                    select max(batch_number) as max_batch from equi_change_equipment_data
                    union
                    select max(batch_number) as max_batch from equi_change_classification
                )
                select max(max_batch) from cte;
                """;
            cmd.CommandText = query;
            if (cmd.ExecuteScalar() is long maxBatch)
            {
                for (int i = 1; i <= maxBatch; i++)
                {
                    Console.WriteLine($"Index: {i}");
                    var uploadTemplatePath = appSettings.EquiChangeTemplatePath;

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
                FROM equi_change_change_request_header t;
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
                FROM equi_change_change_request_notes t;
                """;
            ws = wb.Worksheets.Worksheet("Change Request Header");
            ws.Unprotect();
            using var cmd2 = new SqliteCommand(query, conn);
            using var reader2 = cmd2.ExecuteReader();
            row = 5;
            while (reader2.Read())
            {
                if (!reader2.IsDBNull(0)) ws.Cell(row, "A").Value = reader2.GetString(0);   // usmd_note
                row++;
            };

            // "EQ-Equipment Data" tab
            query = $"""
                SELECT 
                    t.equipment,                -- A
                    t.description_medium,       -- B
                    t.inactive,                 -- C
                    t.object_type,              -- D
                    t.authoriz_group,           -- E
                    t.gross_weight,             -- F
                    t.unit_of_weight,           -- G
                    t.inventory_no,             -- H
                    t.size_dimens,              -- I
                    t.start_up_date,            -- J
                    t.acquisition_Value,        -- K
                    t.currency,                 -- L
                    t.acquistion_date,          -- M
                    t.manufacturer,             -- N
                    t.model_number,             -- O
                    t.manuf_part_no,            -- P
                    t.manuf_serial_number,      -- Q
                    t.manuf_country,            -- R
                    t.construct_year,           -- S
                    t.construct_mth,            -- T
                    t.maint_plant,              -- U
                    t.plant_section,            -- V
                    t.location,                 -- W
                    t.room,                     -- X
                    t.abc_indi,                 -- Y
                    t.work_center,              -- Z
                    t.sort_field,               -- AA
                    t.business_area,            -- AB
                    t.asset,                    -- AC
                    t.sub_number,               -- AD
                    t.cost_center,              -- AE
                    t.wbs_element,              -- AF
                    t.standg_order,             -- AG
                    t.settlement_order,         -- AH
                    t.planning_plant,           -- AI
                    t.planner_group,            -- AJ
                    t.main_work_ctr,            -- AK
                    t.plnt_work_center,         -- AL
                    t.catalog_profile,          -- AM
                    t.functional_loc,           -- AN
                    t.superord_equip,           -- AO
                    t.position,                 -- AP
                    t.tech_ident_no,            -- AQ
                    t.construction_type_ma,     -- AR
                    t.material,                 -- AS
                    t.material_serial_numb,     -- AT
                    t.config_material,          -- AU
                    t.status_profile,           -- AV
                    t.status_of_an_object,      -- AW
                    t.status_without_stsno,     -- AX
                    t.sales_org,                -- AY
                    t.distr_channel,            -- AZ
                    t.division,                 -- BA
                    t.sales_office,             -- BB
                    t.sales_group,              -- BC
                    t.license_no,               -- BD
                    t.begin_guarantee_c,        -- BE
                    t.warranty_end_c,           -- BF
                    t.master_warranty_c,        -- BG
                    t.inherit_warranty_c,       -- BH
                    t.pass_on_warranty_c,       -- BI
                    t.begin_guarantee_v,        -- BJ
                    t.warranty_end_v,           -- BK
                    t.master_warranty_v,        -- BL
                    t.inherit_warranty_v,       -- BM
                    t.pass_on_warr_v,           -- BN
                    t.vendor,                   -- BO
                    t.customer,                 -- BP
                    t.end_customer,             -- BQ
                    t.cur_customer,             -- BR
                    t.operator,                 -- BS
                    t.delivery_date             -- BT          
                FROM equi_change_equipment_data t
                WHERE batch_number = {batch}
                ORDER BY t.functional_loc;
                """;
            ws = wb.Worksheets.Worksheet("EQ-Equipment Data");
            ws.Unprotect();
            using var cmd3 = new SqliteCommand(query, conn);
            using var reader3 = cmd3.ExecuteReader();
            row = 6;
            while (reader3.Read())
            {
                if (!reader3.IsDBNull(0)) ws.Cell(row, "A").Value = reader3.GetString(0);
                if (!reader3.IsDBNull(1)) ws.Cell(row, "B").Value = reader3.GetString(1);
                if (!reader3.IsDBNull(2)) ws.Cell(row, "C").Value = reader3.GetString(2);
                if (!reader3.IsDBNull(3)) ws.Cell(row, "D").Value = reader3.GetString(3);
                if (!reader3.IsDBNull(4)) ws.Cell(row, "E").Value = reader3.GetString(4);
                if (!reader3.IsDBNull(5)) ws.Cell(row, "F").Value = reader3.GetString(5);
                if (!reader3.IsDBNull(6)) ws.Cell(row, "G").Value = reader3.GetString(6);
                if (!reader3.IsDBNull(7)) ws.Cell(row, "H").Value = reader3.GetString(7);
                if (!reader3.IsDBNull(8)) ws.Cell(row, "I").Value = reader3.GetString(8);
                if (!reader3.IsDBNull(9)) ws.Cell(row, "J").Value = reader3.GetString(9);
                if (!reader3.IsDBNull(10)) ws.Cell(row, "K").Value = reader3.GetString(10);
                if (!reader3.IsDBNull(11)) ws.Cell(row, "L").Value = reader3.GetString(11);
                if (!reader3.IsDBNull(12)) ws.Cell(row, "M").Value = reader3.GetString(12);
                if (!reader3.IsDBNull(13)) ws.Cell(row, "N").Value = reader3.GetString(13);
                if (!reader3.IsDBNull(14)) ws.Cell(row, "O").Value = reader3.GetString(14);
                if (!reader3.IsDBNull(15)) ws.Cell(row, "P").Value = reader3.GetString(15);
                if (!reader3.IsDBNull(16)) ws.Cell(row, "Q").Value = reader3.GetString(16);
                if (!reader3.IsDBNull(17)) ws.Cell(row, "R").Value = reader3.GetString(17);
                if (!reader3.IsDBNull(18)) ws.Cell(row, "S").Value = reader3.GetString(18);
                if (!reader3.IsDBNull(19)) ws.Cell(row, "T").Value = reader3.GetString(19);
                if (!reader3.IsDBNull(20)) ws.Cell(row, "U").Value = reader3.GetString(20);
                if (!reader3.IsDBNull(21)) ws.Cell(row, "V").Value = reader3.GetString(21);
                if (!reader3.IsDBNull(22)) ws.Cell(row, "W").Value = reader3.GetString(22);
                if (!reader3.IsDBNull(23)) ws.Cell(row, "X").Value = reader3.GetString(26);
                if (!reader3.IsDBNull(24)) ws.Cell(row, "Y").Value = reader3.GetString(24);
                if (!reader3.IsDBNull(25)) ws.Cell(row, "Z").Value = reader3.GetString(25);
                if (!reader3.IsDBNull(26)) ws.Cell(row, "AA").Value = reader3.GetString(26);
                if (!reader3.IsDBNull(27)) ws.Cell(row, "AB").Value = reader3.GetString(27);
                if (!reader3.IsDBNull(28)) ws.Cell(row, "AC").Value = reader3.GetString(28);
                if (!reader3.IsDBNull(29)) ws.Cell(row, "AD").Value = reader3.GetString(29);
                if (!reader3.IsDBNull(30)) ws.Cell(row, "AE").Value = reader3.GetString(30);
                if (!reader3.IsDBNull(31)) ws.Cell(row, "AF").Value = reader3.GetString(31);
                if (!reader3.IsDBNull(32)) ws.Cell(row, "AG").Value = reader3.GetString(32);
                if (!reader3.IsDBNull(33)) ws.Cell(row, "AH").Value = reader3.GetString(33);
                if (!reader3.IsDBNull(34)) ws.Cell(row, "AI").Value = reader3.GetString(34);
                if (!reader3.IsDBNull(35)) ws.Cell(row, "AJ").Value = reader3.GetString(35);
                if (!reader3.IsDBNull(36)) ws.Cell(row, "AK").Value = reader3.GetString(36);
                if (!reader3.IsDBNull(37)) ws.Cell(row, "AL").Value = reader3.GetString(37);
                if (!reader3.IsDBNull(38)) ws.Cell(row, "AM").Value = reader3.GetString(38);
                if (!reader3.IsDBNull(39)) ws.Cell(row, "AN").Value = reader3.GetString(39);
                if (!reader3.IsDBNull(40)) ws.Cell(row, "AO").Value = reader3.GetString(40);
                if (!reader3.IsDBNull(41)) ws.Cell(row, "AP").Value = reader3.GetString(41);
                if (!reader3.IsDBNull(42)) ws.Cell(row, "AQ").Value = reader3.GetString(42);
                if (!reader3.IsDBNull(43)) ws.Cell(row, "AR").Value = reader3.GetString(43);
                if (!reader3.IsDBNull(44)) ws.Cell(row, "AS").Value = reader3.GetString(44);
                if (!reader3.IsDBNull(45)) ws.Cell(row, "AT").Value = reader3.GetString(45);
                if (!reader3.IsDBNull(46)) ws.Cell(row, "AU").Value = reader3.GetString(46);
                if (!reader3.IsDBNull(47)) ws.Cell(row, "AV").Value = reader3.GetString(47);
                if (!reader3.IsDBNull(48)) ws.Cell(row, "AW").Value = reader3.GetString(48);
                if (!reader3.IsDBNull(49)) ws.Cell(row, "AX").Value = reader3.GetString(49);
                if (!reader3.IsDBNull(50)) ws.Cell(row, "AY").Value = reader3.GetString(50);
                if (!reader3.IsDBNull(51)) ws.Cell(row, "AZ").Value = reader3.GetString(51);
                if (!reader3.IsDBNull(52)) ws.Cell(row, "BA").Value = reader3.GetString(52);
                if (!reader3.IsDBNull(53)) ws.Cell(row, "BB").Value = reader3.GetString(53);
                if (!reader3.IsDBNull(54)) ws.Cell(row, "BC").Value = reader3.GetString(54);
                if (!reader3.IsDBNull(55)) ws.Cell(row, "BD").Value = reader3.GetString(55);
                if (!reader3.IsDBNull(56)) ws.Cell(row, "BE").Value = reader3.GetString(56);
                if (!reader3.IsDBNull(57)) ws.Cell(row, "BF").Value = reader3.GetString(57);
                if (!reader3.IsDBNull(58)) ws.Cell(row, "BG").Value = reader3.GetString(58);
                if (!reader3.IsDBNull(59)) ws.Cell(row, "BH").Value = reader3.GetString(59);
                if (!reader3.IsDBNull(60)) ws.Cell(row, "BI").Value = reader3.GetString(60);
                if (!reader3.IsDBNull(61)) ws.Cell(row, "BJ").Value = reader3.GetString(61);
                if (!reader3.IsDBNull(62)) ws.Cell(row, "BK").Value = reader3.GetString(62);
                if (!reader3.IsDBNull(63)) ws.Cell(row, "BL").Value = reader3.GetString(63);
                if (!reader3.IsDBNull(64)) ws.Cell(row, "BM").Value = reader3.GetString(64);
                if (!reader3.IsDBNull(65)) ws.Cell(row, "BN").Value = reader3.GetString(65);
                if (!reader3.IsDBNull(66)) ws.Cell(row, "BO").Value = reader3.GetString(66);
                if (!reader3.IsDBNull(67)) ws.Cell(row, "BP").Value = reader3.GetString(67);
                if (!reader3.IsDBNull(68)) ws.Cell(row, "BQ").Value = reader3.GetString(68);
                if (!reader3.IsDBNull(69)) ws.Cell(row, "BR").Value = reader3.GetString(69);
                if (!reader3.IsDBNull(70)) ws.Cell(row, "BS").Value = reader3.GetString(70);
                if (!reader3.IsDBNull(71)) ws.Cell(row, "BT").Value = reader3.GetString(71);
                row++;
            };

            // 'EQ-Classification' tab
            query = 
                $"""
                SELECT 
                    t.equipment,
                    t.class,
                    t.characteristics,
                    t.char_value, 
                    t.class_delete_ind
                FROM equi_change_classification t
                WHERE t.batch_number = {batch};
                """;
            ws = wb.Worksheets.Worksheet("EQ-Classification");
            ws.Unprotect();
            using var cmd4 = new SqliteCommand(query, conn);
            using var reader4 = cmd4.ExecuteReader();
            row = 5;
            while (reader4.Read())
            {
                if (!reader4.IsDBNull(0)) ws.Cell(row, "A").Value = reader4.GetString(0);   // Equipment
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
