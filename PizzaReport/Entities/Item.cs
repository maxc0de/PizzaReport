using System;
using System.Text.RegularExpressions;

namespace MilanoExtraReport.BL
{
    public class Item
    {
        private readonly double _amountParts;

        public string Name { get; private set; }
        public string Group { get; private set; }
        public ItemPart Part { get; private set; }
        public double Sum { get; private set; }
        public string Enterprise { get; private set; }

        public double AmountUnit
        {
            get
            {
                if (Part == ItemPart.Full)
                {
                    return _amountParts;
                }
                else
                {
                    return _amountParts / (int)Part;
                }
            }
        }
        public double PriceUnit
        {
            get
            {
                return Math.Round(Sum / AmountUnit);
            }
        }

        public Item(string name, string group, double? amountParts, double? sum, string enterprise)
        {
            if (string.IsNullOrEmpty(name))
            {
                throw new ArgumentException(nameof(name));
            }

            if (string.IsNullOrEmpty(group))
            {
                throw new ArgumentException(nameof(group));
            }

            if (group.Contains("Пицца"))
            {
                Name = SimplifyPizzaName(name);
            }
            else
            {
                Name = name;
            }

            Group = group;
            Sum = (double)sum;
            Enterprise = enterprise;

            Part = GetItemPartFromName(name);

            _amountParts = (double)amountParts;
        }

        private ItemPart GetItemPartFromName(string name)
        {
            if (name.Contains("1/2"))
            {
                return ItemPart.Half;
            }
            else if (name.Contains("1/6"))
            {
                return ItemPart.Sixth;
            }
            else if (name.Contains("1/8"))
            {
                return ItemPart.Eighth;
            }

            return ItemPart.Full;
        }

        private string SimplifyPizzaName(string name)
        {
            Regex regex = new Regex(@".+?(?=\s\d)");
            string pizzaName = regex.Match(name).ToString();

            if (!string.IsNullOrEmpty(pizzaName))
            {
                return pizzaName;
            }

            return name;
        }
    }
}
