using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using SnipperClone.Core;

namespace SnipperClone.Core
{
    public class TableParser
    {
        private static readonly Regex WhitespaceRegex = new Regex(@"\s+", RegexOptions.Compiled);
        private static readonly Regex NumberRegex = new Regex(@"-?\$?(?:\d{1,3}(?:,\d{3})*|\d+)(?:\.\d{1,4})?%?|\(\d+(?:\.\d+)?\)", RegexOptions.Compiled);
        private static readonly Regex TableSeparatorRegex = new Regex(@"[|\t,;]", RegexOptions.Compiled);
        private static readonly Regex LineBreakRegex = new Regex(@"[\r\n]+", RegexOptions.Compiled);
        
        // Enhanced table detection patterns
        private static readonly Regex TableBorderRegex = new Regex(@"^[\s\-_=+|]*$", RegexOptions.Compiled);
        private static readonly Regex AlignmentRegex = new Regex(@"^[\s]*:?[\-_=]+:?[\s]*$", RegexOptions.Compiled);
        
        public TableData ParseTable(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return new TableData();
            }

            try
            {
                System.Diagnostics.Debug.WriteLine("TableParser: Starting enhanced table parsing...");
                
                // Clean and normalize the text
                var cleanedText = EnhancedCleanText(text);
                var lines = SplitIntoLines(cleanedText);
                
                if (lines.Count == 0)
                {
                    return new TableData();
                }

                // Enhanced parsing strategies with priority order
                var strategies = new Func<List<string>, TableData>[]
                {
                    ParseMarkdownTable,      // Highest priority - structured tables
                    ParseTabDelimited,       // High priority - clear delimiters
                    ParsePipeDelimited,      // High priority - clear delimiters
                    ParseCommaDelimited,     // Medium priority - CSV format
                    ParseSemicolonDelimited, // Medium priority - European CSV
                    ParseSpaceDelimited,     // Lower priority - space-based
                    ParseFixedWidth,         // Lower priority - fixed columns
                    ParseStructuredText      // Lowest priority - intelligent text parsing
                };

                TableData bestResult = null;
                var bestScore = 0;
                var strategyResults = new List<(string name, TableData result, int score)>();

                foreach (var strategy in strategies)
                {
                    try
                    {
                        var result = strategy(lines);
                        var score = EvaluateTableQuality(result, lines);
                        
                        strategyResults.Add((strategy.Method.Name, result, score));
                        System.Diagnostics.Debug.WriteLine($"TableParser: Strategy {strategy.Method.Name} scored {score} (rows: {result.Rows.Count}, cols: {result.ColumnCount})");
                        
                        if (score > bestScore)
                        {
                            bestScore = score;
                            bestResult = result;
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"TableParser: Strategy {strategy.Method.Name} failed: {ex.Message}");
                    }
                }

                if (bestResult != null && bestScore > 50) // Minimum quality threshold
                {
                    // Post-process the best result
                    EnhancedPostProcessTable(bestResult);
                    System.Diagnostics.Debug.WriteLine($"TableParser: Successfully parsed table with {bestResult.Rows.Count} rows and {bestResult.ColumnCount} columns using best strategy");
                    return bestResult;
                }

                // Fallback: treat as single column with intelligent row detection
                System.Diagnostics.Debug.WriteLine("TableParser: Using enhanced fallback single-column parsing");
                return CreateIntelligentSingleColumnTable(lines);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"TableParser: Error parsing table: {ex.Message}");
                return new TableData();
            }
        }

        private string EnhancedCleanText(string text)
        {
            if (string.IsNullOrEmpty(text))
                return string.Empty;

            // Remove common OCR artifacts and normalize characters
            text = text.Replace("~", "-")
                      .Replace("¦", "|")
                      .Replace("│", "|")
                      .Replace("┃", "|")
                      .Replace("║", "|")
                      .Replace("┌", "+")
                      .Replace("┐", "+")
                      .Replace("└", "+")
                      .Replace("┘", "+")
                      .Replace("├", "+")
                      .Replace("┤", "+")
                      .Replace("┬", "+")
                      .Replace("┴", "+")
                      .Replace("┼", "+")
                      .Replace("═", "=")
                      .Replace("─", "-");

            // Normalize line endings
            text = text.Replace("\r\n", "\n").Replace("\r", "\n");

            // Clean up excessive whitespace while preserving table structure
            var lines = text.Split('\n');
            var cleanedLines = new List<string>();

            foreach (var line in lines)
            {
                var trimmed = line.Trim();
                if (!string.IsNullOrEmpty(trimmed))
                {
                    // Preserve internal spacing for table structure
                    var cleaned = Regex.Replace(line, @"[ \t]+", " ");
                    cleanedLines.Add(cleaned);
                }
            }

            return string.Join("\n", cleanedLines);
        }

        private List<string> SplitIntoLines(string text)
        {
            return text.Split('\n')
                      .Where(line => !string.IsNullOrWhiteSpace(line))
                      .Select(line => line.Trim())
                      .Where(line => !TableBorderRegex.IsMatch(line)) // Remove border lines
                      .ToList();
        }

        private TableData ParseMarkdownTable(List<string> lines)
        {
            var table = new TableData();
            var headerFound = false;
            var alignmentFound = false;

            for (int i = 0; i < lines.Count; i++)
            {
                var line = lines[i];
                
                // Check for alignment row (markdown table separator)
                if (!headerFound && i < lines.Count - 1 && AlignmentRegex.IsMatch(lines[i + 1]))
                {
                    // This is a header row
                    var headerCells = SplitMarkdownRow(line);
                    table.SetHeaders(headerCells);
                    headerFound = true;
                    alignmentFound = true;
                    i++; // Skip the alignment row
                    continue;
                }

                // Parse data rows
                if (line.Contains("|"))
                {
                    var cells = SplitMarkdownRow(line);
                    if (cells.Length > 0)
                    {
                        table.AddRow(cells);
                    }
                }
            }

            if (table.Rows.Count > 0)
            {
                NormalizeColumnCount(table);
                if (!headerFound)
                {
                    DetectHeaders(table);
                }
            }

            return table;
        }

        private string[] SplitMarkdownRow(string line)
        {
            // Remove leading and trailing pipes
            line = line.Trim();
            if (line.StartsWith("|")) line = line.Substring(1);
            if (line.EndsWith("|")) line = line.Substring(0, line.Length - 1);

            return line.Split('|')
                      .Select(cell => cell.Trim())
                      .ToArray();
        }

        private TableData ParseTabDelimited(List<string> lines)
        {
            var table = new TableData();
            
            foreach (var line in lines)
            {
                if (line.Contains('\t'))
                {
                    var cells = line.Split('\t')
                                   .Select(cell => cell.Trim())
                                   .ToArray();
                    table.AddRow(cells);
                }
            }

            if (table.Rows.Count > 0)
            {
                NormalizeColumnCount(table);
                DetectHeaders(table);
            }

            return table;
        }

        private TableData ParsePipeDelimited(List<string> lines)
        {
            var table = new TableData();
            
            foreach (var line in lines)
            {
                if (line.Contains("|"))
                {
                    var cells = line.Split('|')
                                   .Select(cell => cell.Trim())
                                   .Where(cell => !string.IsNullOrEmpty(cell))
                                   .ToArray();
                    
                    if (cells.Length > 1) // Ensure it's actually a table row
                    {
                        table.AddRow(cells);
                    }
                }
            }

            if (table.Rows.Count > 0)
            {
                NormalizeColumnCount(table);
                DetectHeaders(table);
            }

            return table;
        }

        private TableData ParseCommaDelimited(List<string> lines)
        {
            var table = new TableData();
            
            foreach (var line in lines)
            {
                if (line.Contains(","))
                {
                    var cells = ParseCSVLine(line);
                    if (cells.Length > 1)
                    {
                        table.AddRow(cells);
                    }
                }
            }

            if (table.Rows.Count > 0)
            {
                NormalizeColumnCount(table);
                DetectHeaders(table);
            }

            return table;
        }

        private TableData ParseSemicolonDelimited(List<string> lines)
        {
            var table = new TableData();
            
            foreach (var line in lines)
            {
                if (line.Contains(";"))
                {
                    var cells = line.Split(';')
                                   .Select(cell => cell.Trim())
                                   .ToArray();
                    
                    if (cells.Length > 1)
                    {
                        table.AddRow(cells);
                    }
                }
            }

            if (table.Rows.Count > 0)
            {
                NormalizeColumnCount(table);
                DetectHeaders(table);
            }

            return table;
        }

        private TableData ParseSpaceDelimited(List<string> lines)
        {
            var table = new TableData();
            var columnPositions = DetectColumnPositions(lines);
            
            if (columnPositions.Count < 2)
            {
                return table; // Not enough columns detected
            }

            foreach (var line in lines)
            {
                var cells = ExtractCellsByPosition(line, columnPositions);
                if (cells.Length > 1 && cells.Any(c => !string.IsNullOrWhiteSpace(c)))
                {
                    table.AddRow(cells);
                }
            }

            if (table.Rows.Count > 0)
            {
                NormalizeColumnCount(table);
                DetectHeaders(table);
            }

            return table;
        }

        private TableData ParseFixedWidth(List<string> lines)
        {
            var table = new TableData();
            
            // Analyze character positions to detect column boundaries
            var charFrequency = new Dictionary<int, int>();
            var maxLength = lines.Max(l => l.Length);
            
            // Count spaces at each position
            foreach (var line in lines)
            {
                for (int i = 0; i < line.Length; i++)
                {
                    if (char.IsWhiteSpace(line[i]))
                    {
                        charFrequency[i] = charFrequency.GetValueOrDefault(i, 0) + 1;
                    }
                }
            }
            
            // Find column boundaries (positions with high space frequency)
            var boundaries = new List<int> { 0 };
            var threshold = lines.Count * 0.7; // 70% of lines should have space at this position
            
            for (int i = 1; i < maxLength - 1; i++)
            {
                if (charFrequency.GetValueOrDefault(i, 0) >= threshold)
                {
                    // Check if this is a significant boundary
                    var hasContentBefore = false;
                    var hasContentAfter = false;
                    
                    foreach (var line in lines)
                    {
                        if (i > 0 && i - 1 < line.Length && !char.IsWhiteSpace(line[i - 1]))
                            hasContentBefore = true;
                        if (i + 1 < line.Length && !char.IsWhiteSpace(line[i + 1]))
                            hasContentAfter = true;
                    }
                    
                    if (hasContentBefore && hasContentAfter)
                    {
                        boundaries.Add(i);
                    }
                }
            }
            
            boundaries.Add(maxLength);
            
            if (boundaries.Count < 3) // Need at least 2 columns
            {
                return table;
            }

            // Extract cells based on boundaries
            foreach (var line in lines)
            {
                var cells = new List<string>();
                
                for (int i = 0; i < boundaries.Count - 1; i++)
                {
                    var start = boundaries[i];
                    var end = Math.Min(boundaries[i + 1], line.Length);
                    
                    if (start < line.Length)
                    {
                        var cell = line.Substring(start, end - start).Trim();
                        cells.Add(cell);
                    }
                    else
                    {
                        cells.Add("");
                    }
                }
                
                if (cells.Count > 1 && cells.Any(c => !string.IsNullOrWhiteSpace(c)))
                {
                    table.AddRow(cells.ToArray());
                }
            }

            if (table.Rows.Count > 0)
            {
                NormalizeColumnCount(table);
                DetectHeaders(table);
            }

            return table;
        }

        private TableData ParseStructuredText(List<string> lines)
        {
            var table = new TableData();
            
            // Intelligent text parsing - look for patterns
            var patterns = new[]
            {
                @"(.+?):\s*(.+)",           // Key: Value pairs
                @"(.+?)\s{2,}(.+)",         // Two or more spaces as separator
                @"(.+?)\s+(\d+(?:\.\d+)?)", // Text followed by number
                @"(\d+(?:\.\d+)?)\s+(.+)"  // Number followed by text
            };

            foreach (var line in lines)
            {
                foreach (var pattern in patterns)
                {
                    var match = Regex.Match(line, pattern);
                    if (match.Success)
                    {
                        var cells = new[] { match.Groups[1].Value.Trim(), match.Groups[2].Value.Trim() };
                        table.AddRow(cells);
                        break; // Use first matching pattern
                    }
                }
            }

            if (table.Rows.Count > 0)
            {
                NormalizeColumnCount(table);
                DetectHeaders(table);
            }

            return table;
        }

        private string[] ParseCSVLine(string line)
        {
            var cells = new List<string>();
            var current = "";
            var inQuotes = false;
            
            for (int i = 0; i < line.Length; i++)
            {
                var c = line[i];
                
                if (c == '"')
                {
                    inQuotes = !inQuotes;
                }
                else if (c == ',' && !inQuotes)
                {
                    cells.Add(current.Trim());
                    current = "";
                }
                else
                {
                    current += c;
                }
            }
            
            cells.Add(current.Trim());
            return cells.ToArray();
        }

        private List<int> DetectColumnPositions(List<string> lines)
        {
            var positions = new List<int>();
            var maxLength = lines.Max(l => l.Length);
            
            // Look for consistent word boundaries
            var wordBoundaries = new Dictionary<int, int>();
            
            foreach (var line in lines)
            {
                var words = line.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
                var currentPos = 0;
                
                foreach (var word in words)
                {
                    var wordStart = line.IndexOf(word, currentPos);
                    if (wordStart > 0)
                    {
                        wordBoundaries[wordStart] = wordBoundaries.GetValueOrDefault(wordStart, 0) + 1;
                    }
                    currentPos = wordStart + word.Length;
                }
            }
            
            // Select positions that appear in most lines
            var threshold = lines.Count * 0.5;
            positions.AddRange(wordBoundaries.Where(kvp => kvp.Value >= threshold)
                                           .Select(kvp => kvp.Key)
                                           .OrderBy(p => p));
            
            if (!positions.Contains(0))
                positions.Insert(0, 0);
                
            return positions;
        }

        private string[] ExtractCellsByPosition(string line, List<int> positions)
        {
            var cells = new List<string>();
            
            for (int i = 0; i < positions.Count; i++)
            {
                var start = positions[i];
                var end = i + 1 < positions.Count ? positions[i + 1] : line.Length;
                
                if (start < line.Length)
                {
                    var cell = line.Substring(start, Math.Min(end - start, line.Length - start)).Trim();
                    cells.Add(cell);
                }
                else
                {
                    cells.Add("");
                }
            }
            
            return cells.ToArray();
        }

        private void NormalizeColumnCount(TableData table)
        {
            if (table.Rows.Count == 0)
                return;

            // Find the most common column count
            var columnCounts = table.Rows.GroupBy(row => row.Length)
                                        .OrderByDescending(g => g.Count())
                                        .ToList();

            var targetColumnCount = columnCounts.First().Key;
            table.ColumnCount = targetColumnCount;

            // Normalize all rows to have the same column count
            for (int i = 0; i < table.Rows.Count; i++)
            {
                var row = table.Rows[i];
                
                if (row.Length < targetColumnCount)
                {
                    // Pad with empty cells
                    var newRow = new string[targetColumnCount];
                    Array.Copy(row, newRow, row.Length);
                    for (int j = row.Length; j < targetColumnCount; j++)
                    {
                        newRow[j] = "";
                    }
                    table.Rows[i] = newRow;
                }
                else if (row.Length > targetColumnCount)
                {
                    // Truncate excess columns or merge them
                    var newRow = new string[targetColumnCount];
                    Array.Copy(row, newRow, targetColumnCount - 1);
                    
                    // Merge remaining columns into the last column
                    var mergedCell = string.Join(" ", row.Skip(targetColumnCount - 1));
                    newRow[targetColumnCount - 1] = mergedCell;
                    
                    table.Rows[i] = newRow;
                }
            }
        }

        private void DetectHeaders(TableData table)
        {
            if (table.Rows.Count == 0)
                return;

            var firstRow = table.Rows[0];
            var hasHeader = false;

            // Enhanced header detection
            var headerScore = 0;
            var totalCells = firstRow.Length;
            
            foreach (var cell in firstRow)
            {
                if (!string.IsNullOrWhiteSpace(cell))
                {
                    // Headers typically contain letters and are not pure numbers
                    if (Regex.IsMatch(cell, @"[a-zA-Z]"))
                    {
                        headerScore++;
                    }
                    
                    // Headers often contain specific words
                    if (Regex.IsMatch(cell, @"\b(name|date|amount|total|description|id|number|type|status|category)\b", RegexOptions.IgnoreCase))
                    {
                        headerScore += 2;
                    }
                    
                    // Headers are usually shorter than data
                    if (cell.Length < 20)
                    {
                        headerScore++;
                    }
                    
                    // Headers don't usually contain only numbers
                    if (!NumberRegex.IsMatch(cell.Trim()))
                    {
                        headerScore++;
                    }
                }
            }

            // If most cells in first row look like headers
            hasHeader = headerScore > totalCells * 1.5; // Adjusted threshold

            if (hasHeader && table.Rows.Count > 1)
            {
                table.Headers = firstRow.ToList();
                table.HasHeader = true;
                
                // Remove header row from data rows
                table.Rows.RemoveAt(0);
            }
            else
            {
                table.HasHeader = false;
                table.Headers = Enumerable.Range(1, table.ColumnCount)
                                         .Select(i => $"Column {i}")
                                         .ToList();
            }
        }

        private int EvaluateTableQuality(TableData table, List<string> originalLines)
        {
            if (table == null || table.Rows.Count == 0)
                return 0;

            var score = 0;
            var maxScore = 1000; // Maximum possible score

            try
            {
                // 1. Row consistency (30% of score)
                var rowConsistencyScore = EvaluateRowConsistency(table);
                score += (int)(rowConsistencyScore * 0.30 * maxScore);

                // 2. Column data type consistency (25% of score)
                var dataTypeScore = EvaluateDataTypeConsistency(table);
                score += (int)(dataTypeScore * 0.25 * maxScore);

                // 3. Header quality (20% of score)
                var headerScore = EvaluateHeaderQuality(table);
                score += (int)(headerScore * 0.20 * maxScore);

                // 4. Cell content quality (15% of score)
                var contentScore = EvaluateCellContentQuality(table);
                score += (int)(contentScore * 0.15 * maxScore);

                // 5. Structure integrity (10% of score)
                var structureScore = EvaluateStructureIntegrity(table, originalLines);
                score += (int)(structureScore * 0.10 * maxScore);

                // Bonus points for specific patterns
                score += EvaluateBonusPatterns(table);

                // Penalty for common issues
                score -= EvaluatePenalties(table);

                return Math.Max(0, Math.Min(maxScore, score));
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"TableParser: Error evaluating table quality: {ex.Message}");
                return 0;
            }
        }

        private double EvaluateRowConsistency(TableData table)
        {
            if (table.Rows.Count <= 1) return 0.5;

            var columnCounts = table.Rows.Select(row => row.Length).ToList();
            var mostCommonCount = columnCounts.GroupBy(x => x)
                                            .OrderByDescending(g => g.Count())
                                            .First().Key;

            var consistentRows = columnCounts.Count(c => c == mostCommonCount);
            var consistency = (double)consistentRows / table.Rows.Count;

            // Bonus for having the expected column count
            if (mostCommonCount == table.ColumnCount)
                consistency += 0.1;

            return Math.Min(1.0, consistency);
        }

        private double EvaluateDataTypeConsistency(TableData table)
        {
            if (table.ColumnCount == 0) return 0;

            var consistencyScores = new List<double>();

            for (int col = 0; col < table.ColumnCount; col++)
            {
                var columnValues = table.Rows
                    .Where(row => col < row.Length && !string.IsNullOrWhiteSpace(row[col]))
                    .Select(row => row[col])
                    .ToList();

                if (columnValues.Count == 0)
                {
                    consistencyScores.Add(0);
                    continue;
                }

                var dataTypes = columnValues.Select(GetDataType).ToList();
                var mostCommonType = dataTypes.GroupBy(x => x)
                                             .OrderByDescending(g => g.Count())
                                             .First().Key;

                var typeConsistency = (double)dataTypes.Count(t => t == mostCommonType) / dataTypes.Count;
                consistencyScores.Add(typeConsistency);
            }

            return consistencyScores.Count > 0 ? consistencyScores.Average() : 0;
        }

        private double EvaluateHeaderQuality(TableData table)
        {
            if (!table.HasHeaders || table.Headers == null || table.Headers.Count == 0)
                return 0.3; // Neutral score for no headers

            var score = 0.0;
            var headerCount = table.Headers.Count;

            foreach (var header in table.Headers)
            {
                if (string.IsNullOrWhiteSpace(header))
                {
                    score += 0.1; // Empty headers are poor
                    continue;
                }

                // Check for common header patterns
                if (IsLikelyHeader(header))
                    score += 1.0;
                else if (header.Length > 2 && char.IsLetter(header[0]))
                    score += 0.7;
                else
                    score += 0.3;
            }

            return headerCount > 0 ? score / headerCount : 0;
        }

        private double EvaluateCellContentQuality(TableData table)
        {
            if (table.RowCount == 0 || table.ColumnCount == 0) return 0;

            int numericCells = 0;
            int textCells = 0;
            int emptyCells = 0;
            int totalCells = table.RowCount * table.ColumnCount;

            foreach (var row in table.Rows)
            {
                foreach (var cell in row)
                {
                    if (string.IsNullOrWhiteSpace(cell)) emptyCells++;
                    else if (NumberRegex.IsMatch(cell)) numericCells++;
                    else textCells++;
                }
            }

            double score = 0;
            // Prefer tables with a mix of numeric and text, or mostly numeric/text
            if (numericCells > 0 && textCells > 0) score += 20;
            else if (numericCells > totalCells * 0.7) score += 15; // Mostly numeric
            else if (textCells > totalCells * 0.7) score += 10; // Mostly text

            // Penalize high empty cell ratio, but not too harshly
            double emptyRatio = (double)emptyCells / totalCells;
            score -= emptyRatio * 20; 

            // Bonus for financial or date patterns
            if (HasFinancialPatterns(table)) score += 10;
            if (HasDatePatterns(table)) score += 10;
            if (HasConsistentNumberFormatting(table)) score += 10;

            return Math.Max(0, Math.Min(score, 30)); // Cap score
        }

        private double EvaluateStructureIntegrity(TableData table, List<string> originalLines)
        {
            if (table.RowCount == 0 || table.ColumnCount == 0) return 0;

            double score = 0;
            int maxPossibleScore = 20; // Max score for this category

            // 1. Row count vs original lines (up to 5 points)
            // Considers that some lines might be headers or separators not part of data rows
            double rowMatchRatio = (double)table.RowCount / originalLines.Count;
            if (rowMatchRatio > 0.5 && rowMatchRatio <= 1.1) // Allow slight more rows due to splitting merged cells etc.
            {
                score += 5 * (1 - Math.Abs(1 - rowMatchRatio)); 
            }
            else if (table.RowCount > 0) // Penalize heavily if very different, but give some credit if any rows found
            {
                score += 1;
            }
            
            // 2. Column consistency (up to 10 points)
            if (table.ColumnCount > 0)
            {
                var columnCounts = table.Rows.Select(r => r.Length).ToList();
                if (columnCounts.Any())
                {
                    var mostCommonCount = columnCounts.GroupBy(c => c).OrderByDescending(g => g.Count()).First().Key;
                    var consistentRows = columnCounts.Count(c => c == mostCommonCount);
                    score += 10 * ((double)consistentRows / table.RowCount);
                }
            }

            // 3. Separator consistency (up to 5 points) - This is a more complex check
            // For simplicity here, we're assuming the strategy already picked a consistent separator
            // A more advanced check would re-evaluate based on originalLines and chosen separator
            // For now, let's assume it contributes if a table is found.
            if (table.ColumnCount > 1) // Multi-column tables imply some separator was used
            {
                score += 3; 
            }
            else if (table.ColumnCount == 1 && table.RowCount > 1) // Single column tables are structurally simple
            {
                score += 1;
            }

            return Math.Max(0, Math.Min(score, maxPossibleScore));
        }

        private int EvaluateBonusPatterns(TableData table)
        {
            var bonus = 0;

            // Bonus for financial data patterns
            if (HasFinancialPatterns(table))
                bonus += 50;

            // Bonus for date patterns
            if (HasDatePatterns(table))
                bonus += 30;

            // Bonus for consistent number formatting
            if (HasConsistentNumberFormatting(table))
                bonus += 40;

            // Bonus for proper table size
            if (table.Rows.Count >= 2 && table.Rows.Count <= 50 && table.ColumnCount >= 2 && table.ColumnCount <= 20)
                bonus += 20;

            return bonus;
        }

        private int EvaluatePenalties(TableData table)
        {
            var penalty = 0;

            // Penalty for too many empty cells
            var emptyCellRatio = GetEmptyCellRatio(table);
            if (emptyCellRatio > 0.5)
                penalty += (int)((emptyCellRatio - 0.5) * 200);

            // Penalty for single column tables (usually not real tables)
            if (table.ColumnCount == 1)
                penalty += 100;

            // Penalty for tables with too many or too few rows
            if (table.Rows.Count > 100)
                penalty += 50;
            if (table.Rows.Count == 1)
                penalty += 75;

            return penalty;
        }

        private string GetDataType(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return "empty";

            // Check for numbers (including currency and percentages)
            if (NumberRegex.IsMatch(value))
                return "number";

            // Check for dates
            if (DateTime.TryParse(value, out _))
                return "date";

            // Check for boolean-like values
            if (value.ToLowerInvariant() is "true" or "false" or "yes" or "no" or "y" or "n")
                return "boolean";

            // Check if it's mostly letters
            if (value.Count(char.IsLetter) > value.Length * 0.7)
                return "text";

            return "mixed";
        }

        private bool IsLikelyHeader(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return false;

            var lowerText = text.ToLowerInvariant();
            
            // Common header words
            var headerWords = new[]
            {
                "name", "date", "amount", "total", "description", "type", "status", "id", "number",
                "account", "balance", "debit", "credit", "reference", "category", "item", "quantity",
                "price", "value", "code", "title", "subject", "period", "year", "month", "day"
            };

            return headerWords.Any(word => lowerText.Contains(word)) ||
                   (text.Length > 2 && text.Length < 30 && char.IsUpper(text[0]));
        }

        private double EvaluateSeparatorConsistency(List<string> lines)
        {
            var separatorCounts = new Dictionary<char, int>
            {
                { '\t', 0 }, { '|', 0 }, { ',', 0 }, { ';', 0 }
            };

            foreach (var line in lines)
            {
                foreach (var sep in separatorCounts.Keys.ToList())
                {
                    separatorCounts[sep] += line.Count(c => c == sep);
                }
            }

            var totalSeparators = separatorCounts.Values.Sum();
            if (totalSeparators == 0) return 0;

            var maxCount = separatorCounts.Values.Max();
            return (double)maxCount / totalSeparators;
        }

        private bool HasFinancialPatterns(TableData table)
        {
            var financialKeywords = new[] { "$", "€", "£", "¥", "₹", "%", "total", "amount", "balance", "debit", "credit" };
            
            return table.Rows.Any(row => 
                row.Any(cell => 
                    financialKeywords.Any(keyword => 
                        cell.ToLowerInvariant().Contains(keyword))));
        }

        private bool HasDatePatterns(TableData table)
        {
            return table.Rows.Any(row => 
                row.Any(cell => 
                    DateTime.TryParse(cell, out _) || 
                    Regex.IsMatch(cell, @"\d{1,2}[/-]\d{1,2}[/-]\d{2,4}")));
        }

        private bool HasConsistentNumberFormatting(TableData table)
        {
            var numberColumns = new List<int>();
            
            for (int col = 0; col < table.ColumnCount; col++)
            {
                var columnValues = table.Rows
                    .Where(row => col < row.Length)
                    .Select(row => row[col])
                    .Where(cell => !string.IsNullOrWhiteSpace(cell))
                    .ToList();

                if (columnValues.Count > 0 && columnValues.Count(NumberRegex.IsMatch) > columnValues.Count * 0.7)
                {
                    numberColumns.Add(col);
                }
            }

            return numberColumns.Count > 0;
        }

        private double GetEmptyCellRatio(TableData table)
        {
            var totalCells = table.Rows.Sum(row => row.Length);
            var emptyCells = table.Rows.Sum(row => row.Count(string.IsNullOrWhiteSpace));
            
            return totalCells > 0 ? (double)emptyCells / totalCells : 0;
        }

        private void EnhancedPostProcessTable(TableData table)
        {
            if (table == null || table.Rows.Count == 0)
                return;

            // Clean up cell contents
            for (int i = 0; i < table.Rows.Count; i++)
            {
                for (int j = 0; j < table.Rows[i].Length; j++)
                {
                    var cell = table.Rows[i][j];
                    if (!string.IsNullOrEmpty(cell))
                    {
                        // Remove extra whitespace
                        cell = WhitespaceRegex.Replace(cell, " ").Trim();
                        
                        // Clean up common OCR artifacts in numbers
                        if (NumberRegex.IsMatch(cell))
                        {
                            cell = cell.Replace("O", "0")  // Common OCR mistake
                                      .Replace("l", "1")   // Common OCR mistake
                                      .Replace("S", "5")   // Common OCR mistake
                                      .Replace("B", "8");  // Common OCR mistake
                        }
                        
                        // Clean up text
                        cell = cell.Replace("~", "-")
                                  .Replace("¦", "|");
                        
                        table.Rows[i][j] = cell;
                    }
                }
            }

            // Update headers if they exist
            if (table.HasHeader && table.Headers != null)
            {
                for (int i = 0; i < table.Headers.Count; i++)
                {
                    if (!string.IsNullOrEmpty(table.Headers[i]))
                    {
                        table.Headers[i] = WhitespaceRegex.Replace(table.Headers[i], " ").Trim();
                    }
                }
            }

            // Remove completely empty rows
            table.Rows.RemoveAll(row => row.All(cell => string.IsNullOrWhiteSpace(cell)));

            // Recalculate column count after cleanup
            if (table.Rows.Count > 0)
            {
                table.ColumnCount = table.Rows.Max(row => row.Length);
            }
        }

        private TableData CreateIntelligentSingleColumnTable(List<string> lines)
        {
            var table = new TableData
            {
                ColumnCount = 1,
                HasHeader = false,
                Headers = new List<string> { "Data" }
            };

            // Intelligent row detection - group related lines
            var currentGroup = new List<string>();
            
            foreach (var line in lines)
            {
                var trimmed = line.Trim();
                
                // Check if this line starts a new logical row
                if (IsNewRowIndicator(trimmed, currentGroup))
                {
                    if (currentGroup.Count > 0)
                    {
                        var combinedText = string.Join(" ", currentGroup).Trim();
                        if (!string.IsNullOrWhiteSpace(combinedText))
                        {
                            table.AddRow(new[] { combinedText });
                        }
                        currentGroup.Clear();
                    }
                }
                
                currentGroup.Add(trimmed);
            }
            
            // Add the last group
            if (currentGroup.Count > 0)
            {
                var combinedText = string.Join(" ", currentGroup).Trim();
                if (!string.IsNullOrWhiteSpace(combinedText))
                {
                    table.AddRow(new[] { combinedText });
                }
            }

            return table;
        }

        private bool IsNewRowIndicator(string line, List<string> currentGroup)
        {
            if (currentGroup.Count == 0)
                return true;

            // Indicators that this might be a new row:
            // 1. Line starts with a number
            if (Regex.IsMatch(line, @"^\d+[\.\)]\s"))
                return true;
                
            // 2. Line starts with a bullet point
            if (Regex.IsMatch(line, @"^[\-\*\•]\s"))
                return true;
                
            // 3. Line starts with a capital letter after a line that ends with punctuation
            var lastLine = currentGroup.LastOrDefault();
            if (!string.IsNullOrEmpty(lastLine) && 
                Regex.IsMatch(lastLine, @"[\.!?]$") && 
                Regex.IsMatch(line, @"^[A-Z]"))
                return true;
                
            // 4. Significant indentation change
            var currentIndent = line.Length - line.TrimStart().Length;
            var lastIndent = lastLine?.Length - lastLine?.TrimStart().Length ?? 0;
            if (Math.Abs(currentIndent - lastIndent) > 2)
                return true;

            return false;
        }
    }

    public class TableStructure
    {
        public int ColumnCount { get; set; } = 1;
        public string SeparatorType { get; set; } = "spaces";
        public bool HasHeader { get; set; } = false;
    }
} 