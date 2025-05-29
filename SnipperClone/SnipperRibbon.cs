using System;
using System.IO;
using System.Reflection;

namespace SnipperClone
{
    /// <summary>
    /// Helper class for loading ribbon XML resources
    /// </summary>
    public static class SnipperRibbon
    {
        /// <summary>
        /// Gets the ribbon XML content from embedded resources
        /// </summary>
        /// <returns>The ribbon XML as a string</returns>
        public static string GetRibbonXml()
        {
            try
            {
                var assembly = Assembly.GetExecutingAssembly();
                using (var stream = assembly.GetManifestResourceStream("SnipperClone.SnipperRibbon.xml"))
                {
                    if (stream != null)
                    {
                        using (var reader = new StreamReader(stream))
                        {
                            return reader.ReadToEnd();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading ribbon XML: {ex.Message}");
            }

            // Return null if resource not found - Connect.cs will use fallback
            return null;
        }
    }
} 