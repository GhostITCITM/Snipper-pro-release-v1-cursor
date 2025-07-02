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

            // Remove common currency/percentage symbols and whitespaces
            value = value.Replace("$", string.Empty)
                         .Replace("%", string.Empty)
                         .Replace(" ", string.Empty);

            bool hasDot = value.Contains('.');
            bool hasComma = value.Contains(',');

            if (hasDot && hasComma)
            {
                // Determine which comes first to guess thousands/decimal separator
                int dotIndex = value.IndexOf('.');
                int commaIndex = value.IndexOf(',');

                if (dotIndex < commaIndex)
                {
                    // Format like 1.234,56 -> dot thousands, comma decimal
                    value = value.Replace(".", string.Empty)
                                 .Replace(',', '.');
                }
                else
                {
                    // Format like 1,234.56 -> comma thousands, dot decimal
                    value = value.Replace(",", string.Empty);
                }
            }
            else if (hasComma && !hasDot)
            {
                // Only commas present - treat as decimal separator
                value = value.Replace(',', '.');
            }
            else
            {
                // Only dot or none - remove stray commas
                value = value.Replace(",", string.Empty);
            }

            // Handle accounting format with parentheses
            value = value.Replace("(", "-").Replace(")", string.Empty);

            return double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out number);
        }
    }
}
