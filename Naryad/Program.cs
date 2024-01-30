using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2010.Word.Drawing;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;

namespace Naryad {

    public class Month_U {
        
        public int count_of_days;
        public double percent_need;
        public double percent_of_day;
        public int count_of_pzkv = 0;
        public int corr_count_of_pzkv;
        public double percent_fact;

        public void GetPercentOfDay() {
            percent_of_day = percent_need / count_of_days;
        }

        public void CreateTableList(List<int> vk_naryad_need) {

            int[,] table_of_april = new int[vk_naryad_need.Count, this.count_of_days];

            for (int i = 0; i < vk_naryad_need.Count; i++) {

                for (int j = 0; j < this.count_of_days; j++) {

                    table_of_april[i, j] = Convert.ToInt32(this.percent_of_day * vk_naryad_need[i]);

                    this.count_of_pzkv += table_of_april[i, j];

                    //Console.Write(table_of_april[i, j]);

                }

                //Console.Write("\n");
            }
        }

        public void DoCor(int naryad_fact, double naryad) {

            if(naryad_fact > naryad && (this.count_of_pzkv/naryad) > this.percent_need) {

                count_of_pzkv--;

            }

        }

    }

    internal class Program {

        static void Main(string[] args) {

            string file_path = "C:\\Users\\Максим\\source\\Naryad.xlsx";

            var upload_workbook = new XLWorkbook(file_path);
            var worksheet = upload_workbook.Worksheet(1);

            var naryad = worksheet.Cell("C11").Value.GetNumber();

            var vk_naryad_need = new List<int>();
            var import_rows = worksheet.Range("C14:C44");

            foreach (var cell in import_rows.Cells()) {

                vk_naryad_need.Add(Convert.ToInt32(cell.CachedValue.GetNumber()));

            }

            Month_U april = new Month_U() {
                count_of_days = Convert.ToInt32(worksheet.Cell("D2").Value.GetNumber()),
                percent_need = worksheet.Cell("C2").Value.GetNumber()
            };

            Month_U may = new Month_U() {
                count_of_days = Convert.ToInt32(worksheet.Cell("D3").Value.GetNumber()),
                percent_need = worksheet.Cell("C3").Value.GetNumber()
            };

            Month_U june = new Month_U() {
                count_of_days = Convert.ToInt32(worksheet.Cell("D4").Value.GetNumber()),
                percent_need = worksheet.Cell("C4").Value.GetNumber()
            };

            Month_U july = new Month_U() {
                count_of_days = Convert.ToInt32(worksheet.Cell("D5").Value.GetNumber()),
                percent_need = worksheet.Cell("C5").Value.GetNumber()
            };

            april.GetPercentOfDay();
            may.GetPercentOfDay();
            june.GetPercentOfDay();
            july.GetPercentOfDay();

            april.CreateTableList(vk_naryad_need);
            may.CreateTableList(vk_naryad_need);
            june.CreateTableList(vk_naryad_need);
            july.CreateTableList(vk_naryad_need);

            int naryad_fact = april.count_of_pzkv + may.count_of_pzkv + june.count_of_pzkv + july.count_of_pzkv;

            while(naryad_fact != naryad) {

                april.DoCor(naryad_fact, naryad);
                may.DoCor(naryad_fact, naryad);
                june.DoCor(naryad_fact, naryad);
                july.DoCor(naryad_fact, naryad);

                naryad_fact = april.count_of_pzkv + may.count_of_pzkv + june.count_of_pzkv + july.count_of_pzkv;

            }


            //Console.WriteLine(april.count_of_pzkv);
            //Console.WriteLine(may.count_of_pzkv);
            //Console.WriteLine(june.count_of_pzkv);
            //Console.WriteLine(july.count_of_pzkv);

            Console.WriteLine(april.count_of_pzkv / naryad);
            Console.WriteLine(may.count_of_pzkv / naryad);
            Console.WriteLine(june.count_of_pzkv / naryad);
            Console.WriteLine(july.count_of_pzkv / naryad);

            Console.WriteLine(naryad_fact);

            Console.ReadLine();

        }
    }
}
