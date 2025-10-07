using System;
using System.IO;
using System.Reflection;

namespace FocaExcelExport
{
    internal static class AssemblyResolver
    {
        private static bool _initialized;
        
        public static void Init()
        {
            if (_initialized) return;
            _initialized = true;
            AppDomain.CurrentDomain.AssemblyResolve += ResolveFromPluginFolder;
        }

        private static Assembly ResolveFromPluginFolder(object sender, ResolveEventArgs args)
        {
            try
            {
                var name = new AssemblyName(args.Name).Name + ".dll";
                var baseDir = Path.GetDirectoryName(typeof(AssemblyResolver).Assembly.Location);
                var probePaths = new[]
                {
                    baseDir,
                    Path.Combine(baseDir ?? "", "lib"),
                };
                
                foreach (var dir in probePaths)
                {
                    if (string.IsNullOrEmpty(dir)) continue;
                    var candidate = Path.Combine(dir, name);
                    if (File.Exists(candidate)) 
                        return Assembly.LoadFrom(candidate);
                }
            }
            catch { /* swallow and let runtime continue */ }
            return null;
        }
    }
}