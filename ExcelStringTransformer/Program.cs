using System;
using System.Linq;
using ClosedXML.Excel;

namespace ExcelStringTransformer
{
    class Program
    {
        static void Main(string[] args)
        {
            // Excelファイルのパスを指定
            string filePath = @"aaa.xlsx";

            // 新しいExcelワークブックを開く
            var workbook = new XLWorkbook(filePath);

            // "WORD"シートの値を取得
            var wordSheet = workbook.Worksheet("WORD");
            var tarValues = wordSheet.Column("B").CellsUsed().Select(cell => cell.GetString()).ToList();
            var repValues = wordSheet.Column("C").CellsUsed().Select(cell => cell.GetString()).ToList();

            // すべてのシートに対して処理を行う
            foreach (var worksheet in workbook.Worksheets)
            {
                // "WORD"シート以外の処理を行う
                if (worksheet.Name != "WORD")
                {
                    for (int i = 0; i < tarValues.Count(); i++)
                    {
                        string targetValue = tarValues[i];

                        foreach (var cell in worksheet.CellsUsed())
                        {
                            if (cell.HasFormula) continue; // 数式を持つセルはスキップ

                            if (cell.Value.ToString() == targetValue)
                            {
                                cell.FormulaA1 = "=" + repValues[i]; // 文字列を置き換える
                            }
                        }
                    }

                }
            }

            // 変更を保存
            workbook.SaveAs(filePath);
        }
    }
}
