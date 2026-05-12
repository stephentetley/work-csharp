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

namespace SptUtils.ReadAsFitted {

    public class ReadTestSheet
    {
        IXLWorksheet sheet;

        public ReadTestSheet(IXLWorksheet ws)
        {
            sheet = ws;
        }

        public bool isTestSheet()
        {
            var value = sheet.Cell("F1").GetString();
            return (value == "TEST SHEET");

        }

        public void readTest()
        {
            readHeaders();
        }

        private void readHeaders()
        {
            var siteName = sheet.Cell("B3").GetString();
            var dbOrPanelNumber = sheet.Cell("E3").GetString();
            var date = sheet.Cell("J3").GetString();

            Console.WriteLine(siteName);
            Console.WriteLine(dbOrPanelNumber);
            Console.WriteLine(date);
        }


    }
}
