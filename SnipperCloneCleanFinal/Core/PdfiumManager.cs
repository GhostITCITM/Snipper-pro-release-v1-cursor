using System;
using System.IO;
using System.Runtime.InteropServices;
using SnipperCloneCleanFinal.Infrastructure;

namespace SnipperCloneCleanFinal.Core
{
    /// <summary>
    /// Manages PDFium native library loading with multiple fallback strategies
    /// </summary>
    public static class PdfiumManager
    {
        private static bool _isInitialized = false;
        private static bool _isLoaded = false;
        private static string _loadedPath = null;

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr LoadLibrary(string dllToLoad);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool SetDllDirectory(string lpPathName);

        [DllImport("kernel32.dll")]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool FreeLibrary(IntPtr hModule);

        /// <summary>
        /// Initialize PDFium with multiple loading strategies
        /// </summary>
        public static bool Initialize()
        {
            if (_isInitialized)
                return _isLoaded;

            _isInitialized = true;
            Logger.Info("Initializing PDFium manager...");

            try
            {
                // Strategy 1: Check if already loaded
                if (IsPdfiumAlreadyLoaded())
                {
                    Logger.Info("PDFium already loaded in process");
                    _isLoaded = true;
                    return true;
                }

                // Strategy 2: Load from application directory
                if (LoadFromApplicationDirectory())
                {
                    _isLoaded = true;
                    return true;
                }

                // Strategy 3: Load from current working directory
                if (LoadFromCurrentDirectory())
                {
                    _isLoaded = true;
                    return true;
                }

                // Strategy 4: Load from system paths
                if (LoadFromSystemPaths())
                {
                    _isLoaded = true;
                    return true;
                }

                // Strategy 5: Copy and load appropriate architecture
                if (CopyAndLoadArchitectureSpecific())
                {
                    _isLoaded = true;
                    return true;
                }

                Logger.Error("All PDFium loading strategies failed");
                return false;
            }
            catch (Exception ex)
            {
                Logger.Error($"PDFium initialization failed: {ex.Message}", ex);
                return false;
            }
        }

        private static bool IsPdfiumAlreadyLoaded()
        {
            try
            {
                var handle = GetModuleHandle("pdfium.dll");
                if (handle != IntPtr.Zero)
                {
                    Logger.Info("PDFium already loaded in current process");
                    return true;
                }
            }
            catch (Exception ex)
            {
                Logger.Info($"Error checking if PDFium loaded: {ex.Message}");
            }
            return false;
        }

        private static bool LoadFromApplicationDirectory()
        {
            try
            {
                var appDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                
                // Try default pdfium.dll first
                var pdfiumPath = Path.Combine(appDir, "pdfium.dll");
                Logger.Info($"Trying to load PDFium from application directory: {pdfiumPath}");
                
                if (File.Exists(pdfiumPath))
                {
                    SetDllDirectory(appDir);
                    var handle = LoadLibrary(pdfiumPath);
                    if (handle != IntPtr.Zero)
                    {
                        Logger.Info($"Successfully loaded PDFium from: {pdfiumPath}");
                        _loadedPath = pdfiumPath;
                        return true;
                    }
                    else
                    {
                        var error = Marshal.GetLastWin32Error();
                        Logger.Error($"Failed to load PDFium from {pdfiumPath}, error: {error}");
                        
                        // If 64-bit failed and we're on 64-bit, try x86 version
                        if (Environment.Is64BitProcess)
                        {
                            var x86Path = Path.Combine(appDir, "pdfium_x86.dll");
                            if (File.Exists(x86Path))
                            {
                                Logger.Info($"Trying x86 version: {x86Path}");
                                handle = LoadLibrary(x86Path);
                                if (handle != IntPtr.Zero)
                                {
                                    Logger.Info($"Successfully loaded x86 PDFium from: {x86Path}");
                                    _loadedPath = x86Path;
                                    return true;
                                }
                            }
                        }
                    }
                }
                else
                {
                    Logger.Info($"PDFium not found at: {pdfiumPath}");
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"Error loading PDFium from application directory: {ex.Message}", ex);
            }
            return false;
        }

        private static bool LoadFromCurrentDirectory()
        {
            try
            {
                var currentDir = Directory.GetCurrentDirectory();
                var pdfiumPath = Path.Combine(currentDir, "pdfium.dll");
                
                Logger.Info($"Trying to load PDFium from current directory: {pdfiumPath}");
                
                if (File.Exists(pdfiumPath))
                {
                    SetDllDirectory(currentDir);
                    var handle = LoadLibrary(pdfiumPath);
                    if (handle != IntPtr.Zero)
                    {
                        Logger.Info($"Successfully loaded PDFium from: {pdfiumPath}");
                        _loadedPath = pdfiumPath;
                        return true;
                    }
                }
                else
                {
                    Logger.Info($"PDFium not found at: {pdfiumPath}");
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"Error loading PDFium from current directory: {ex.Message}", ex);
            }
            return false;
        }

        private static bool LoadFromSystemPaths()
        {
            try
            {
                Logger.Info("Trying to load PDFium from system paths");
                var handle = LoadLibrary("pdfium.dll");
                if (handle != IntPtr.Zero)
                {
                    Logger.Info("Successfully loaded PDFium from system paths");
                    _loadedPath = "system";
                    return true;
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"Error loading PDFium from system paths: {ex.Message}", ex);
            }
            return false;
        }

        private static bool CopyAndLoadArchitectureSpecific()
        {
            try
            {
                // Determine if we're running x86 or x64
                bool is64Bit = Environment.Is64BitProcess;
                Logger.Info($"Process architecture: {(is64Bit ? "x64" : "x86")}");

                var appDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                var targetPath = Path.Combine(appDir, "pdfium.dll");

                // Find the project root (go up directories until we find packages folder)
                var currentDir = appDir;
                string packagesDir = null;
                
                for (int i = 0; i < 5; i++) // Search up to 5 levels up
                {
                    var testPackagesDir = Path.Combine(currentDir, "packages");
                    if (Directory.Exists(testPackagesDir))
                    {
                        packagesDir = testPackagesDir;
                        break;
                    }
                    var parent = Directory.GetParent(currentDir);
                    if (parent == null) break;
                    currentDir = parent.FullName;
                }

                if (packagesDir == null)
                {
                    Logger.Error("Could not find packages directory");
                    return false;
                }

                // Determine source path based on architecture
                string sourcePath;
                if (is64Bit)
                {
                    sourcePath = Path.Combine(packagesDir, "PdfiumViewer.Native.x86_64.no_v8-no_xfa.2018.4.8.256", "Build", "x64", "pdfium.dll");
                }
                else
                {
                    sourcePath = Path.Combine(packagesDir, "PdfiumViewer.Native.x86.no_v8-no_xfa.2018.4.8.256", "Build", "x86", "pdfium.dll");
                }

                Logger.Info($"Trying to copy PDFium from: {sourcePath} to: {targetPath}");

                if (File.Exists(sourcePath))
                {
                    // Copy the appropriate DLL
                    File.Copy(sourcePath, targetPath, true);
                    Logger.Info($"Successfully copied {(is64Bit ? "x64" : "x86")} PDFium DLL");

                    // Try to load it
                    SetDllDirectory(appDir);
                    var handle = LoadLibrary(targetPath);
                    if (handle != IntPtr.Zero)
                    {
                        Logger.Info($"Successfully loaded PDFium from: {targetPath}");
                        _loadedPath = targetPath;
                        return true;
                    }
                    else
                    {
                        var error = Marshal.GetLastWin32Error();
                        Logger.Error($"Failed to load copied PDFium, error: {error}");
                    }
                }
                else
                {
                    Logger.Error($"Source PDFium not found at: {sourcePath}");
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"Error copying and loading architecture-specific PDFium: {ex.Message}", ex);
            }
            return false;
        }

        /// <summary>
        /// Test if PDFium functions can be called
        /// </summary>
        public static bool TestPdfiumFunctions()
        {
            if (!_isLoaded)
            {
                Logger.Info("PDFium not loaded, cannot test functions");
                return false;
            }

            try
            {
                // Try to call a simple PDFium function to verify it's working
                // This will throw an exception if the DLL isn't properly loaded
                FPDF_GetLastError();
                Logger.Info("PDFium functions test: SUCCESS");
                return true;
            }
            catch (Exception ex)
            {
                Logger.Error($"PDFium functions test failed: {ex.Message}", ex);
                return false;
            }
        }

        public static string GetLoadedPath()
        {
            return _loadedPath ?? "Not loaded";
        }

        public static bool IsLoaded => _isLoaded;

        // Test function to verify PDFium is working
        [DllImport("pdfium.dll", EntryPoint = "FPDF_GetLastError")]
        private static extern uint FPDF_GetLastError();
    }
} 