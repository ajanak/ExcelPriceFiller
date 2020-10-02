using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;


namespace ExceplPriceFilter.ConsoleApp
{
    class Program
    {
        private const string equipmentLabel = "Оборудование";
        private const string materialsLabel = "Изделия и материалы";
        
        private const string specIdLabel = "Тип, марка, обозначение документа, опросного листа";
        private const string intemsCountLabel = "Кол.";
        private const string priceLabel = "Цена за ед";
        private const string totalPriceLabel = "Цена итого";
        private const string mountingPriceLabel = "Цена монтажа за ед";
        private const string totalMouningPriceLabel = "Цена монтажа итого";

        private const string totalEquipmentPriceLabel = "Итого Обор:";
        private const string totalMaterialsPriceLabel = "Итого Мат:";

        private static readonly List<string> allLabels = new List<string>
        {
            equipmentLabel, materialsLabel,
            specIdLabel, intemsCountLabel, priceLabel, totalPriceLabel,
            mountingPriceLabel, totalMouningPriceLabel, totalEquipmentPriceLabel, totalMaterialsPriceLabel
        };
        
        static void Main(string[] args)
        {
            var filePath = args[0];
            
            var workbook = ReadWorkBook(filePath);


            var sheetNumber = 2;
            var controlPoints = GetSheetControlPoints(workbook, sheetNumber, allLabels);
            
            //todo:: validate controlPoints

            var priceBySpec = GetPriceList().ToDictionary(x => x.SpecId);
            
            ProcessSheetControlPointSet(workbook, sheetNumber, controlPoints, priceBySpec);
            
            WriteWorkBook("text.xlsx", workbook);
        }

        private static IWorkbook ReadWorkBook(string filePath)
        {
            using var file = new FileStream(
                filePath, 
                FileMode.Open, 
                FileAccess.Read);

            return new XSSFWorkbook(file);
        }
        
        private static void WriteWorkBook(string filePath, IWorkbook workbook)
        {
            using var file = new FileStream(
                filePath, 
                FileMode.Create, 
                FileAccess.Write);

            workbook.Write(file);
        }

        private static void ProcessSheetControlPointSet(
            IWorkbook workBook, 
            int sheetNumber,
            Dictionary<string, Point> controlPoints,
            Dictionary<string, PriceInfo> priceOfSpec)
        {
            var sheet = workBook.GetSheetAt(sheetNumber);
            
            for (var rowNumber = controlPoints[equipmentLabel].RowNumber + 1;
                rowNumber < controlPoints[totalEquipmentPriceLabel].RowNumber;
                ++rowNumber)
            {
                var row = sheet.GetRow(rowNumber);
                
                var specCell = row.GetCell(controlPoints[specIdLabel].CellNumber);
                var itemCountCell = row.GetCell(controlPoints[intemsCountLabel].CellNumber);
                var priceForOneItemCell = row.GetCell(controlPoints[priceLabel].CellNumber);
                var totalPriceCell = row.GetCell(controlPoints[totalPriceLabel].CellNumber);
                var priceForOneItemMountingCell = row.GetCell(controlPoints[mountingPriceLabel].CellNumber);
                var totalPriceForMountingCell = row.GetCell(controlPoints[totalMouningPriceLabel].CellNumber);

                if (specCell.CellType != CellType.String)
                {
                    //todo::log значение в колонке не строкового типа
                }

                if (priceOfSpec.TryGetValue(specCell.StringCellValue, out var priceInfo))
                {
                    priceForOneItemCell.SetCellValue(priceInfo.Price);
                    priceForOneItemMountingCell.SetCellValue(priceInfo.PriceForMounting);
                    totalPriceCell.SetCellValue(itemCountCell.NumericCellValue * priceInfo.Price);
                    totalPriceForMountingCell.SetCellValue(itemCountCell.NumericCellValue * priceInfo.PriceForMounting);
                }
                else
                {
                    //todo::log не найдена цена
                }
            }
            
            for (var rowNumber = controlPoints[materialsLabel].RowNumber + 1;
                rowNumber < controlPoints[totalMaterialsPriceLabel].RowNumber;
                ++rowNumber)
            {
                var row = sheet.GetRow(rowNumber);
                
                var specCell = row.GetCell(controlPoints[specIdLabel].CellNumber);
                var itemCountCell = row.GetCell(controlPoints[intemsCountLabel].CellNumber);
                var priceForOneItemCell = row.GetCell(controlPoints[priceLabel].CellNumber);
                var totalPriceCell = row.GetCell(controlPoints[totalPriceLabel].CellNumber);
                var priceForOneItemMountingCell = row.GetCell(controlPoints[mountingPriceLabel].CellNumber);
                var totalPriceForMountingCell = row.GetCell(controlPoints[totalMouningPriceLabel].CellNumber);

                if (specCell.CellType != CellType.String)
                {
                    //todo::log значение в колонке не строкового типа
                }

                if (priceOfSpec.TryGetValue(specCell.StringCellValue, out var priceInfo))
                {
                    priceForOneItemCell.SetCellValue(priceInfo.Price);
                    priceForOneItemMountingCell.SetCellValue(priceInfo.PriceForMounting);
                    totalPriceCell.SetCellValue(itemCountCell.NumericCellValue * priceInfo.Price);
                    totalPriceForMountingCell.SetCellValue(itemCountCell.NumericCellValue * priceInfo.PriceForMounting);
                }
                else
                {
                    //todo::log не найдена цена
                }
            }
        }

        private static IEnumerable<PriceInfo> GetPriceList()
        {
            yield return new PriceInfo("С2000-БИ", 1111, 2222);
        }

        private static Dictionary<string, Point> GetSheetControlPoints(
            IWorkbook workBook,
            int sheetNumber,
            IReadOnlyList<string> requiredControlPointValues)
        {
            var result = new Dictionary<string, Point>();

            var sheetName = workBook.GetSheetName(sheetNumber);
            var sheet = workBook.GetSheet(sheetName);

            for (var rowNum = sheet.FirstRowNum; rowNum <= sheet.LastRowNum; ++rowNum)
            {
                var row = sheet.GetRow(rowNum);
                if (row == null)
                {
                    continue;
                }

                for (var cellNum = row.FirstCellNum; cellNum <= row.LastCellNum; ++cellNum)
                {
                    var cell = row.GetCell(cellNum);
                    if (cell == null)
                    {
                        continue;
                    }

                    if (cell.CellType == CellType.String && requiredControlPointValues.Contains(cell.StringCellValue))
                    {
                        if (result.ContainsKey(cell.StringCellValue))
                        {
                            //todo:: логирование дублируется значение опорной точки, это ошибка, прерываем обработку
                        }

                        result[cell.StringCellValue] = new Point(rowNum, cellNum);
                    }
                }
            }

            return result;
        }
        

        private class Point
        {
            public int RowNumber { get; }
            
            public int CellNumber { get; }

            public Point(int rowNumber, int cellNumber)
            {
                RowNumber = rowNumber;
                CellNumber = cellNumber;
            }
        }

        private class PriceInfo
        {
            public string SpecId { get; }

            public double Price { get;  }
            
            public double PriceForMounting { get; }

            public PriceInfo(string specId, double price, double priceForMounting)
            {
                SpecId = specId;
                Price = price;
                PriceForMounting = priceForMounting;
            }
        }
    }
}