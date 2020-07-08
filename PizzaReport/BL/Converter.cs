using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace MilanoExtraReport.BL
{
    public static class Converter
    {
        private static Application _excel;

        public static event Action<int> Read;
        public static event Action<int> Write;
        public static event Action<string> Converted;
        public static event Action<Exception> Error;

        public static void Convert(string fileName)
        {
            string newFileName = fileName.Replace(".xlsx", "_new.xlsx");

            try
            {
                IEnumerable<Item> items = GetListItem(fileName);

                if (items != null)
                {
                    CreateExcel(newFileName, AddSameTypePizza(items));
                }
            }
            catch(Exception ex)
            {
                Error(ex);
                return;
            }
            finally
            {
                if(_excel != null)
                {
                    _excel.Quit();
                }
            }
        }

        private static IEnumerable<Item> GetListItem(string fileName)
        {
            _excel = new Application();

            Workbook objWorkBook = _excel.Workbooks.Open(fileName, 0, false, 5, "", "", false, XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            Worksheet objWorkSheet = (Worksheet)objWorkBook.Sheets[1];

            List<Item> pizzaItems = new List<Item>();
            string enterprise = string.Empty;
            string group = null;

            int amountRows = objWorkSheet.UsedRange.Rows.Count;

            for (int i = 1; i <= amountRows; i++)
            {
                string enterpriseColumn   = objWorkSheet.Cells[i, 1].Value as string;
                string groupColumn        = objWorkSheet.Cells[i, 2].Value as string;
                string nameColumn         = objWorkSheet.Cells[i, 3].Value as string;
                double? amountPartsColumn = objWorkSheet.Cells[i, 4].Value as double?;
                double? sumColumn         = objWorkSheet.Cells[i, 5].Value as double?;

                if (TryParseEnterprise(enterpriseColumn, out string parsedEnterprise))
                {
                    enterprise = parsedEnterprise;
                }

                if (TryParsePizzaItemGroup(groupColumn, out string parsedGroup))
                {
                    group = parsedGroup;
                }

                if (TryParsePizzaItemName(nameColumn, out string parcedName) && group != null)
                {
                    pizzaItems.Add(new Item(parcedName, group, amountPartsColumn, sumColumn, enterprise));
                }

                Read(amountRows);
            }

            _excel.Quit();

            return pizzaItems;
        }

        private static bool TryParseEnterprise(string enterprise, out string parsedEnterprise)
        {
            parsedEnterprise = null;

            if (enterprise != null && enterprise.Contains("_") && !enterprise.Contains("всего"))
            {
                parsedEnterprise = enterprise;
                return true;
            }

            return false;
        }

        private static bool TryParsePizzaItemGroup(string group, out string parsedGroup)
        {
            parsedGroup = null;

            if (group != null && !group.Contains("всего") && !group.Contains("Группа блюда"))
            {
                parsedGroup = group;
                return true;
            }

            return false;
        }

        private static bool TryParsePizzaItemName(string name, out string parcedName)
        {
            parcedName = null;

            if (name == null || name.Contains("Блюдо"))
            {
                return false;
            }

            parcedName = name;
            return true;
        }

        private static IEnumerable<Item> AddSameTypePizza(IEnumerable<Item> pizzas)
        {
            List<Item> items = new List<Item>();

            var enterprises = pizzas.GroupBy(p => p.Enterprise);

            foreach(var enterprise in enterprises)
            {
                IEnumerable<Item> eight = enterprise.Where(p => p.Group.Contains("8") && p.Group.Contains("Пицца")).ToList();
                IEnumerable<Item> six = enterprise.Where(p => p.Group.Contains("6") && p.Group.Contains("Пицца")).ToList();

                items.AddRange(eight.GroupBy(p => p.Name).Select(p => new Item(p.First().Name, "Пицца 8", p.Sum(a => a.AmountUnit), p.Sum(b => b.Sum), p.First().Enterprise)));
                items.AddRange(six.GroupBy(p => p.Name).Select(p => new Item(p.First().Name, "Пицца 6", p.Sum(a => a.AmountUnit), p.Sum(b => b.Sum), p.First().Enterprise)));

                items.AddRange(enterprise.Except(eight).Except(six));
            }

            return items;
        }

        private static void CreateExcel(string fileName, IEnumerable<Item> pizzas)
        {
            _excel = new Application();

            Workbook objWorkBook = _excel.Workbooks.Add(Type.Missing);
            Worksheet objWorkSheet = (Worksheet)objWorkBook.Sheets[1];

            objWorkSheet.Cells[1, 1] = "Подразделение";
            objWorkSheet.Cells[1, 2] = "Группа";
            objWorkSheet.Cells[1, 3] = "Наименование";
            objWorkSheet.Cells[1, 4] = "Цена продажи";
            objWorkSheet.Cells[1, 5] = "Количество";
            objWorkSheet.Cells[1, 6] = "Сумма";

            int i = 2;
            foreach (Item pizzaItem in pizzas)
            {
                objWorkSheet.Cells[i, 1] = pizzaItem.Enterprise;
                objWorkSheet.Cells[i, 2] = pizzaItem.Group;
                objWorkSheet.Cells[i, 3] = pizzaItem.Name;
                objWorkSheet.Cells[i, 4] = pizzaItem.PriceUnit;
                objWorkSheet.Cells[i, 5] = pizzaItem.AmountUnit;
                objWorkSheet.Cells[i, 6] = pizzaItem.Sum;
                i++;
                Write(pizzas.Count());
            }

            objWorkBook.SaveAs(fileName);
            Converted(fileName);
        }
    }
}
