using System.Text.RegularExpressions;

namespace SnipperCloneCleanFinal.Core
{
    internal static class FormulaParser
    {
        private static readonly Regex Rx = new(@"(?:DS\.|SnipperPro\.Connect\.)(?:\w+)\(\"(?<id>[^\"]+)\"\)", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        public static bool TryGetId(string formula, out string id)
        {
            id = null;
            var m = Rx.Match(formula ?? string.Empty);
            if (m.Success)
            {
                id = m.Groups["id"].Value;
                return true;
            }
            return false;
        }
    }
}
