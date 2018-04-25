using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.Eventing;
using System.Linq;
using System.Reflection;

namespace SimpleXlsxDiff
{
    class Program
    {
        static void Main(string[] args)
        {
            Trace.Listeners.Add(new TextWriterTraceListener(Console.Out));

            var assemblies = new Dictionary<string,Assembly>();
            var executingAssembly = Assembly.GetExecutingAssembly();
            var resources = executingAssembly.GetManifestResourceNames().Where(val => val.ToLower().EndsWith(".dll"));

            foreach (var resource in resources)
            {
                using (var stream = executingAssembly.GetManifestResourceStream(resource))
                {
                    if(stream == null)
                        continue;

                    var buffer = new byte[stream.Length];
                    stream.Read(buffer, 0, buffer.Length);
                    try
                    {
                        assemblies.Add(resource.ToLower(),Assembly.Load(buffer));
                    }
                    catch (Exception e)
                    {
                        Trace.TraceWarning(e.Message);
                    }
                }
            }

            AppDomain.CurrentDomain.AssemblyResolve += (sender, eventArgs) =>
            {
                var assemblyName = new AssemblyName(eventArgs.Name);

                var path = assemblyName.Name.ToLower() + ".dll";
                Assembly assembly;
                assemblies.TryGetValue(path, out assembly);
                return assembly;
            };

            MainProcess.Execute(args);
        }
    }
}
