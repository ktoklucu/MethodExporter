using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using ClosedXML.Excel;

namespace MethodExporter
{
    public class MethodInfo
    {
        public string Folder { get; set; }
        public string File { get; set; }
        public string Namespace { get; set; }
        public string ClassChain { get; set; }
        public string Signature { get; set; }
        public string FullName => $"{Namespace}.{ClassChain}.{Signature}";
    }

    public class InvocationInfo
    {
        public string CallerFullName { get; set; }
        public string InvokedFullName { get; set; }
    }

    public static class MethodExtractor
    {
        public static IEnumerable<string> GetAllCsFiles(string rootPath) =>
            Directory.EnumerateFiles(rootPath, "*.cs", SearchOption.AllDirectories);

        public static List<MethodInfo> ExtractMethods(
            SyntaxTree tree,
            string filePath,
            string rootFolder)
        {
            var root = tree.GetCompilationUnitRoot();
            var ns = root.DescendantNodes()
                         .OfType<NamespaceDeclarationSyntax>()
                         .FirstOrDefault()?.Name.ToString() ?? string.Empty;

            string GetClassChain(SyntaxNode node) =>
                string.Join('.',
                    node.Ancestors()
                        .OfType<TypeDeclarationSyntax>()
                        .Reverse()
                        .Select(t => t.Identifier.Text)
                );

            return root.DescendantNodes()
                       .OfType<MethodDeclarationSyntax>()
                       .Select(m =>
                       {
                           var returnType = m.ReturnType.ToString();
                           var name = m.Identifier.Text;
                           var parameters = string.Join(
                               ", ",
                               m.ParameterList.Parameters
                                .Select(p => $"{p.Type} {p.Identifier}"));

                           return new MethodInfo
                           {
                               Folder = new DirectoryInfo(rootFolder).Name,
                               File = Path.GetRelativePath(rootFolder, filePath),
                               Namespace = ns,
                               ClassChain = GetClassChain(m),
                               Signature = $"{returnType} {name}({parameters})"
                           };
                       })
                       .ToList();
        }

        public static List<InvocationInfo> ExtractInvocations(
            SyntaxTree tree,
            string filePath,
            CSharpCompilation compilation)
        {
            var root = tree.GetCompilationUnitRoot();
            var model = compilation.GetSemanticModel(tree);

            string GetClassChain(SyntaxNode node) =>
                string.Join('.',
                    node.Ancestors()
                        .OfType<TypeDeclarationSyntax>()
                        .Reverse()
                        .Select(t => t.Identifier.Text)
                );

            var ns = root.DescendantNodes()
                         .OfType<NamespaceDeclarationSyntax>()
                         .FirstOrDefault()?.Name.ToString() ?? string.Empty;

            var invs = new List<InvocationInfo>();

            foreach (var invoke in root.DescendantNodes().OfType<InvocationExpressionSyntax>())
            {
                var methodDecl = invoke.Ancestors()
                                       .OfType<MethodDeclarationSyntax>()
                                       .FirstOrDefault();
                if (methodDecl == null) continue;

                // Reconstruct MethodInfo.FullName exactly
                var returnType = methodDecl.ReturnType.ToString();
                var methodName = methodDecl.Identifier.Text;
                var parameters = string.Join(
                    ", ",
                    methodDecl.ParameterList.Parameters
                    .Select(p => $"{p.Type} {p.Identifier}"));
                var signature = $"{returnType} {methodName}({parameters})";
                var classChain = GetClassChain(methodDecl);
                var callerFull = $"{ns}.{classChain}.{signature}";

                // Determine invoked name
                var exprText = invoke.Expression.ToString();
                var symInfo = model.GetSymbolInfo(invoke);
                var invokedSymbol = symInfo.Symbol as IMethodSymbol
                                   ?? symInfo.CandidateSymbols.OfType<IMethodSymbol>().FirstOrDefault();

                ITypeSymbol receiverType = null;
                if (invoke.Expression is MemberAccessExpressionSyntax ma)
                {
                    var recvInfo = model.GetSymbolInfo(ma.Expression).Symbol;
                    switch (recvInfo)
                    {
                        case ILocalSymbol loc: receiverType = loc.Type; break;
                        case IFieldSymbol fld: receiverType = fld.Type; break;
                        case IPropertySymbol prop: receiverType = prop.Type; break;
                    }
                }

                bool isServiceCall = receiverType != null
                 && receiverType.Name.EndsWith("Service");

                string? invokedName;
                if (exprText.Contains(".Instance.") || isServiceCall)
                {
                    invokedName = exprText;
                }
                else
                {
                    invokedName = null;
                }

                invs.Add(new InvocationInfo
                {
                    CallerFullName = callerFull,
                    InvokedFullName = invokedName
                });

                Console.WriteLine($"Invoked Name: {invokedName}");
            }

            return [.. invs.Where(s => s.InvokedFullName != null)];
        }

        public static List<InvocationInfo> ComputeTransitiveInvocations(
            List<InvocationInfo> direct)
        {
            var graph = direct.GroupBy(x => x.CallerFullName)
                .ToDictionary(
                    g => g.Key,
                    g => g.Select(x => x.InvokedFullName).Distinct().ToList()
                );

            var closure = new HashSet<(string caller, string callee)>(
                direct.Select(x => (x.CallerFullName, x.InvokedFullName))
            );

            foreach (var caller in graph.Keys)
            {
                var seen = new HashSet<string>();
                var stack = new Stack<string>(graph[caller]);
                while (stack.Count > 0)
                {
                    var callee = stack.Pop();
                    if (!seen.Add(callee)) continue;
                    closure.Add((caller, callee));
                    if (graph.TryGetValue(callee, out var nexts))
                        foreach (var nxt in nexts)
                            stack.Push(nxt);
                }
            }

            return closure.Select(t => new InvocationInfo
            {
                CallerFullName = t.caller,
                InvokedFullName = t.callee
            }).ToList();
        }

        public static void WriteToExcel(
            IEnumerable<MethodInfo> methods,
            IEnumerable<InvocationInfo> invocations,
            string outFolder)
        {
            var path = Path.Combine(outFolder, $"MethodList.xlsx");
            using var wb = new XLWorkbook();

            var ws1 = wb.Worksheets.Add("Methods");
            ws1.Cell(1, 1).Value = "Folder";
            ws1.Cell(1, 2).Value = "File";
            ws1.Cell(1, 3).Value = "Namespace";
            ws1.Cell(1, 4).Value = "ClassChain";
            ws1.Cell(1, 5).Value = "Signature";
            ws1.Cell(1, 6).Value = "FullName";
            int row = 2;
            foreach (var m in methods)
            {
                ws1.Cell(row, 1).Value = m.Folder;
                ws1.Cell(row, 2).Value = m.File;
                ws1.Cell(row, 3).Value = m.Namespace;
                ws1.Cell(row, 4).Value = m.ClassChain;
                ws1.Cell(row, 5).Value = m.Signature;
                ws1.Cell(row, 6).Value = m.FullName;
                row++;
            }
            ws1.Columns().AdjustToContents();

            var ws2 = wb.Worksheets.Add("Invocations");
            ws2.Cell(1, 1).Value = "Caller";
            ws2.Cell(1, 2).Value = "Invoked";
            row = 2;
            foreach (var inv in invocations)
            {
                ws2.Cell(row, 1).Value = inv.CallerFullName;
                ws2.Cell(row, 2).Value = inv.InvokedFullName;
                row++;
            }
            ws2.Columns().AdjustToContents();

            wb.SaveAs(path);
            Console.WriteLine($"Excel kaydedildi: {path}");
        }

        public static void AppendInstanceCallsColumn(
            IEnumerable<InvocationInfo> invocations,
            string excelFilePath)
        {
            using var wb = new XLWorkbook(excelFilePath);
            var ws = wb.Worksheet("Methods");

            var lastCol = ws.LastColumnUsed().ColumnNumber();
            var newCol = lastCol + 1;
            ws.Cell(1, newCol).Value = "InstanceCalls";

            var instMap = invocations
                .Where(i => i.InvokedFullName.Contains(".Instance."))
                .GroupBy(i => i.CallerFullName)
                .ToDictionary(
                    g => g.Key,
                    g => g.Select(x => x.InvokedFullName).Distinct().ToList()
                );

            var lastRow = ws.LastRowUsed().RowNumber();
            for (int r = 2; r <= lastRow; r++)
            {
                var callerFull = ws.Cell(r, 6).GetString();
                if (instMap.TryGetValue(callerFull, out var calls))
                    ws.Cell(r, newCol).Value = string.Join(Environment.NewLine, calls);
            }

            wb.Save();
        }
    }

    internal class Program
    {
        private static MetadataReference? RefFrom(Type t)
        {
            var loc = t.Assembly.Location;
            return string.IsNullOrWhiteSpace(loc) ? null : MetadataReference.CreateFromFile(loc);
        }


        private static void Main(string[] args)
        {
            Console.Write("Analiz edilecek klasör veya .cs dosyası yolu: ");
            var input = Console.ReadLine()?.Trim();
            if (string.IsNullOrEmpty(input))
            {
                Console.WriteLine("Geçersiz yol, çıkılıyor.");
                return;
            }

            string rootFolder;
            List<string> files;
            if (File.Exists(input) && Path.GetExtension(input).Equals(".cs", StringComparison.OrdinalIgnoreCase))
            {
                rootFolder = Path.GetDirectoryName(input)!;
                files = new List<string> { input };
                Console.WriteLine($"Tek dosya modu: {Path.GetFileName(input)}");
            }
            else if (Directory.Exists(input))
            {
                rootFolder = input;
                files = MethodExporter.MethodExtractor.GetAllCsFiles(rootFolder).ToList();
                Console.WriteLine($"{files.Count} .cs dosyası bulundu.");
            }
            else
            {
                Console.WriteLine("Ne dosya ne klasör, çıkılıyor.");
                return;
            }

            var trees = files.Select(f => CSharpSyntaxTree.ParseText(File.ReadAllText(f))).ToList();

            var refs = new[]
            {
                    RefFrom(typeof(object)),
                    RefFrom(typeof(Enumerable)),
                    RefFrom(typeof(XLWorkbook)),
                    RefFrom(typeof(CSharpSyntaxTree))
            }.Where(r => r != null).ToList();

            var compilation = CSharpCompilation.Create(
                "TmpCompilation",
                trees,
                refs,
                new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary)
            );

            Console.Write("Excel çıkışı için klasör (boş=masaüstü): ");
            var outDir = Console.ReadLine()?.Trim();
            var excelFolder = !string.IsNullOrEmpty(outDir) && Directory.Exists(outDir)
                              ? outDir
                              : Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            var decls = new List<MethodInfo>();
            var dirs = new List<InvocationInfo>();
            for (int i = 0; i < files.Count; i++)
            {
                decls.AddRange(MethodExtractor.ExtractMethods(trees[i], files[i], rootFolder));
                dirs.AddRange(MethodExtractor.ExtractInvocations(trees[i], files[i], compilation));
            }

            var allInvs = MethodExtractor.ComputeTransitiveInvocations(dirs);

            var excelPath = Path.Combine(excelFolder, "MethodList.xlsx");
            MethodExtractor.WriteToExcel(decls, allInvs, excelFolder);
            MethodExtractor.AppendInstanceCallsColumn(allInvs, excelPath);
        }
    }
}
