using System.Text;
using System.Text.Json;
using DocxExtraction;

return Run(args);

static int Run(string[] args)
{
    const string usage = """
        Extract structured data from a Word DOCX using TX Text Control.

        Usage:
          dotnet run -- [input.docx] [-o output.json] [--compact]

        With no input, extracts the included Samples/customer-order.docx.
        JSON goes to stdout unless --output/-o is supplied. Diagnostics go to stderr.
        Exports document structure and values without business mappings or numeric parsing.
        Exit codes: 0 success (see JSON warnings), 1 processing error, 2 usage error.
        """;

    if (args.Contains("--help") || args.Contains("-h"))
    {
        Console.WriteLine(usage);
        return 0;
    }

    try
    {
        string? input = null;
        string? output = null;
        var compact = false;
        for (var i = 0; i < args.Length; i++)
        {
            string NextValue()
            {
                if (i + 1 >= args.Length || args[i + 1].StartsWith('-'))
                    throw new ArgumentException($"Missing value after {args[i]}.");
                return args[++i];
            }
            switch (args[i])
            {
                case "--output": case "-o": output = NextValue(); break;
                case "--compact": compact = true; break;
                default:
                    if (args[i].StartsWith('-')) throw new ArgumentException($"Unknown option: {args[i]}");
                    if (input is not null) throw new ArgumentException("Only one input DOCX is supported per invocation.");
                    input = args[i];
                    break;
            }
        }
        input ??= Path.Combine(AppContext.BaseDirectory, "Samples", "customer-order.docx");
        input = Path.GetFullPath(input);
        if (!input.EndsWith(".docx", StringComparison.OrdinalIgnoreCase)) throw new ArgumentException("Input must be a .docx file.");
        if (output is not null && string.Equals(input, Path.GetFullPath(output), StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("Output must be different from the input document.");

        var result = new DocumentExtractor().Extract(input);
        var json = JsonSerializer.Serialize(result, new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            WriteIndented = !compact
        });
        if (output is null)
        {
            Console.OutputEncoding = new UTF8Encoding(false);
            Console.WriteLine(json);
        }
        else
        {
            File.WriteAllText(output, json + Environment.NewLine, new UTF8Encoding(false));
            Console.Error.WriteLine($"Extracted {result.Tables.Count} tables, {result.ContentControls.Count} content controls, and {result.LegacyFormFields.Count} legacy fields to {Path.GetFullPath(output)}");
        }
        return 0;
    }
    catch (ArgumentException ex)
    {
        Console.Error.WriteLine(ex.Message);
        Console.Error.WriteLine("Use --help for usage.");
        return 2;
    }
    catch (Exception ex)
    {
        Console.Error.WriteLine($"Extraction failed: {ex.Message}");
        return 1;
    }
}
