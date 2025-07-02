using System;
using System.Globalization;

namespace SnipperCloneCleanFinal.Core
{
    internal static class NumberHelper
    {
        public static bool TryParseFlexible(string input, out double number)
        {
            number = 0;
            if (string.IsNullOrWhiteSpace(input))
                return false;

            var value = input.Trim();
            value = value.Replace("$", "").Replace("%", "");

            bool hasDot = value.Contains('.');
            bool hasComma = value.Contains(',');

            if (hasComma && !hasDot)
            {
                // Assume comma is decimal separator
                value = value.Replace(',', '.');
            }
            else
            {
                // Remove thousands separators
                value = value.Replace(",", "");
            }

            value = value.Replace("(", "-").Replace(")", "");
            value = value.Trim();

            return double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out number);
        }
    }
}
