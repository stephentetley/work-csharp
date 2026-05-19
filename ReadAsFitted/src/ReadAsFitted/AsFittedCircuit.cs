// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at

// http://www.apache.org/licenses/LICENSE-2.0

// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.



using System.Text.Json;

namespace SptUtils.ReadAsFitted 
{
    public record AsFittedCircuit
    (
        string FileName, 
        string TabName,
        int ChecklistYear,
        string SiteName,
        string DbOrPanelNumber,
        string HeaderTestDate,
        string SheetNumber,
        string AibRef,
        string TpOrSp,
        string Location,
        string DbOrPanelIncomerDetails,
        string CableNum,
        string FedFrom,
        string CircuitRefAndPhase,
        string CircuitDescription,
        string CircuitType,
        string CableType,
        string InstallationMethod,
        string CableLength,
        string NumOfCoresCSA,
        string CircuitBreakerOrFuseRating,
        string CircuitBreakerBSAndTypeNum,
        string CircuitBreakerManufacturerAndRefNum,
        string RCDManufacturerAndType,
        string Load,
        string RatingKW,
        string FullLoadCurrentA,
        string CircuitVoltageV,
        string CircuitCurrentA,
        string TestDate,
        string Comments
    )
    {
        public void WriteJson(Utf8JsonWriter writer)
        {
            writer.WriteStartObject();
            writer.WriteString("file_name", FileName);
            writer.WriteString("sheet_name", TabName);
            writer.WriteNumber("checklist_year", ChecklistYear);
            writer.WriteString("site_name", SiteName);
            writer.WriteString("db_or_panel_number", DbOrPanelNumber);
            writer.WriteString("header_test_date", HeaderTestDate);
            writer.WriteString("sheet_number", SheetNumber);
            writer.WriteString("aib_ref", AibRef);
            writer.WriteString("tp_or_sp", TpOrSp);
            writer.WriteString("location", Location);
            writer.WriteString("db_or_panel_incomer_details", DbOrPanelIncomerDetails);
            writer.WriteString("cable_num", CableNum);
            writer.WriteString("fed_from", FedFrom);
            writer.WriteString("circuit_ref_and_phase", CircuitRefAndPhase);
            writer.WriteString("circuit_description", CircuitDescription);
            writer.WriteString("circuit_type", CircuitType);
            writer.WriteString("cable_type", CableType);
            writer.WriteString("installation_method", InstallationMethod);
            writer.WriteString("cable_length", CableLength);
            writer.WriteString("num_of_cores_csa", NumOfCoresCSA);
            writer.WriteString("circuit_breaker_or_fuse_rating", CircuitBreakerOrFuseRating);
            writer.WriteString("circuit_breaker_bs_and_type_num", CircuitBreakerBSAndTypeNum);
            writer.WriteString("circuit_breaker_manufacturer_and_ref_num", CircuitBreakerManufacturerAndRefNum);
            writer.WriteString("rcd_manufacturer_and_type", RCDManufacturerAndType);
            writer.WriteString("load", Load);
            writer.WriteString("rating_kw", RatingKW);
            writer.WriteString("full_load_current_a", FullLoadCurrentA);
            writer.WriteString("circuit_voltage_v", CircuitVoltageV);
            writer.WriteString("circuit_current_a", CircuitCurrentA);
            writer.WriteString("test_date", TestDate);
            writer.WriteString("comments", Comments);
            writer.WriteEndObject();
        }
    };
    
}
