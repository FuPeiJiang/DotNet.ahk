using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.Emit;
using System.IO.Pipes;
using System.Text;

namespace CompilerWrapper
{
    public static class Program
    {
        public static List<MetadataReference>? cachedReferences = null;

        public static async Task Main()
        {
            while (true)
            {
                var server = new NamedPipeServerStream(
                    "CSharp-Compiler-Wrapper",
                    PipeDirection.InOut,
                    NamedPipeServerStream.MaxAllowedServerInstances,
                    PipeTransmissionMode.Byte,
                    PipeOptions.Asynchronous
                );

                await server.WaitForConnectionAsync();

                // Handle client on a new task/thread
                _ = Task.Run(() => HandleClient(server));
            }
        }

        static void HandleClient(NamedPipeServerStream pipe)
        {
            using (pipe)
            {
                using var reader = new BinaryReader(pipe, Encoding.UTF8, leaveOpen: true);
                using var writer = new BinaryWriter(pipe, Encoding.UTF8, leaveOpen: true);

                string assemblyName = reader.ReadString();
                string externalReferences = reader.ReadString();
                string code = reader.ReadString();

                if (assemblyName == "")
                {
                    assemblyName = Path.GetRandomFileName();
                }
                SyntaxTree syntaxTree = CSharpSyntaxTree.ParseText(code);

                if (cachedReferences == null)
                {
                    string version = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription[5..]; // ".NET 10.0.1"
                    string tfm = AppContext.TargetFrameworkName![21..]; // ".NETCoreApp,Version=v10.0"

                    cachedReferences = [
                        ..Directory.EnumerateFiles(
                            $"C:\\Program Files\\dotnet\\packs\\Microsoft.NETCore.App.Ref\\{version}\\ref\\net{tfm}",
                            "*.dll",
                            SearchOption.TopDirectoryOnly
                        ).Select(p => MetadataReference.CreateFromFile(p)),
                        ..Directory.EnumerateFiles(
                            $"C:\\Program Files\\dotnet\\packs\\Microsoft.WindowsDesktop.App.Ref\\{version}\\ref\\net{tfm}",
                            "*.dll",
                            SearchOption.TopDirectoryOnly
                        ).Select(p => MetadataReference.CreateFromFile(p)),
                    ];

                }
                IEnumerable<MetadataReference> references;

                if (string.IsNullOrEmpty(externalReferences))
                {
                    references = cachedReferences;
                }
                else
                {
                    var mergedReferences = cachedReferences.ToList(); // copy
                    mergedReferences.AddRange(externalReferences.Split("|").Select(p => MetadataReference.CreateFromFile(p)));
                    references = mergedReferences;
                }

                CSharpCompilation compilation = CSharpCompilation.Create(
                    assemblyName,
                    syntaxTrees: [syntaxTree],
                    references: references,
                    options: new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary));

                using MemoryStream ms = new();
                EmitResult result = compilation.Emit(ms);
                Console.WriteLine("finished Compiling");
                if (result.Success)
                {
                    ms.Position = 0;
                    writer.Write("success");
                    writer.Write7BitEncodedInt64(ms.Length);
                    ms.CopyTo(pipe);
                }
                else
                {
                    writer.Write("failure");
                    Console.WriteLine("Compilation failed!");
                    IEnumerable<Diagnostic> failures = result.Diagnostics.Where(diagnostic =>
                        diagnostic.IsWarningAsError ||
                        diagnostic.Severity == DiagnosticSeverity.Error);

                    writer.Write(String.Join("\n", failures.Select(diagnostic => $"\t{diagnostic.Id}: {diagnostic.GetMessage()}")));
                }

            }
        }

    }

}