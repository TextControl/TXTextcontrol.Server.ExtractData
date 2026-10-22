using TXTextControl;
using TXTextControl.DocumentServer.Fields;

namespace DocxExtraction;

/// <summary>Reusable extraction service. Every call owns and disposes its TX instance.</summary>
public sealed class DocumentExtractor
{
    public ExtractionResult Extract(string path) =>
        Extract(File.ReadAllBytes(path), Path.GetFileName(path));

    public ExtractionResult Extract(byte[] docx, string sourceFile)
    {
        var warnings = new List<string>();
        var controls = new List<ContentControlData>();
        var legacy = new List<LegacyFormFieldData>();
        var applicationFields = new List<ApplicationFieldData>();
        var tables = new List<TableData>();
        var sections = new List<SectionData>();

        // MSWord retains native Word SDT metadata for the DocumentServer field adapters.
        using var tx = Load(docx, ApplicationFieldFormat.MSWord);
        ReadWordFields(tx.ApplicationFields.Cast<ApplicationField>(), "body", controls, legacy, applicationFields, warnings);
        foreach (Section section in tx.Sections)
        foreach (HeaderFooter hf in section.HeadersAndFooters)
            ReadWordFields(hf.ApplicationFields.Cast<ApplicationField>(), Location(section, hf), controls, legacy, applicationFields, warnings);

        ReadTables(tx.Tables.Cast<Table>(), "body", tables);
        foreach (Section section in tx.Sections)
        {
            tx.Selection.Start = section.Start - 1; // TX element starts are one-based; selections are zero-based.
            tx.Selection.Length = section.Length;
            var text = tx.Selection.Text;
            var headers = new List<HeaderFooterData>();
            foreach (HeaderFooter hf in section.HeadersAndFooters)
            {
                hf.Selection.Start = 0;
                hf.Selection.Length = -1;
                headers.Add(new(hf.Type.ToString(), hf.ConnectedToPrevious, hf.Selection.Text));
                ReadTables(hf.Tables.Cast<Table>(), Location(section, hf), tables);
            }
            sections.Add(new(section.Number, section.Start, section.Length, text,
                section.Format.BreakKind.ToString(), section.Format.Landscape, headers));
        }

        var settings = tx.DocumentSettings;
        var custom = new Dictionary<string, object?>();
        if (settings.UserDefinedDocumentProperties is { } properties)
            foreach (string key in properties.Names)
                custom[key] = properties[key];
        var metadata = new DocumentMetadata(settings.DocumentTitle, settings.Author, settings.DocumentSubject,
            settings.DocumentKeywords ?? [], settings.CreationDate == DateTime.MinValue ? null : settings.CreationDate,
            settings.CreatorApplication, custom);

        return new(sourceFile, metadata,
            controls, legacy, applicationFields, tables, sections, warnings.Distinct().ToArray());
    }

    private static ServerTextControl Load(byte[] docx, ApplicationFieldFormat format)
    {
        var tx = new ServerTextControl();
        try
        {
            tx.Create();
            tx.Load(docx, BinaryStreamType.WordprocessingML, new LoadSettings { ApplicationFieldFormat = format });
            return tx;
        }
        catch { tx.Dispose(); throw; }
    }

    private static string Location(Section section, HeaderFooter hf) => $"section:{section.Number}/{hf.Type}";

    private static void ReadWordFields(IEnumerable<ApplicationField> fields, string location,
        List<ContentControlData> controls, List<LegacyFormFieldData> legacy,
        List<ApplicationFieldData> other, List<string> warnings)
    {
        foreach (var field in fields)
        {
            try
            {
                if (field.TypeName.StartsWith("SDT", StringComparison.OrdinalIgnoreCase))
                {
                    var control = ContentControlFieldAdapter.CreateContentControl(field);
                    var choices = control switch
                    {
                        DropDownListContentControl dropdown => dropdown.ListItems.Select(i => new ChoiceData(i.DisplayText, i.Value)).ToArray(),
                        ComboBoxContentControl combo => combo.ListItems.Select(i => new ChoiceData(i.DisplayText, i.Value)).ToArray(),
                        _ => Array.Empty<ChoiceData>()
                    };
                    controls.Add(new(location, control.Id, control.Title, control.Tag, control.GetType().Name,
                        field.Text, control is CheckBoxContentControl check ? check.Checked : null,
                        control is DateContentControl date && !string.IsNullOrWhiteSpace(field.Text) ? DateOnly.FromDateTime(date.Date) : null, choices));
                    continue;
                }
                switch (field.TypeName.ToUpperInvariant())
                {
                    case "FORMTEXT":
                        var text = new FormText(field);
                        legacy.Add(new(location, text.Name, "Text", field.Text, null, []));
                        break;
                    case "FORMCHECKBOX":
                        var check = new FormCheckBox(field);
                        legacy.Add(new(location, check.Name, "CheckBox", field.Text, check.Checked, []));
                        break;
                    case "FORMDROPDOWN":
                        var dropdown = new FormDropDown(field);
                        legacy.Add(new(location, dropdown.Name, "DropDown", field.Text, null, dropdown.ListEntries.ToArray()));
                        break;
                    default:
                        other.Add(new(location, field.TypeName, field.Name, field.Text, field.Parameters ?? []));
                        break;
                }
            }
            catch (Exception ex) when (ex is ArgumentException or InvalidOperationException or FormatException or System.Xml.XmlException)
            {
                warnings.Add($"Could not adapt {field.TypeName} at {location}, character {field.Start}: {ex.Message}");
                other.Add(new(location, field.TypeName, field.Name, field.Text, field.Parameters ?? []));
            }
        }
    }

    private static void ReadTables(IEnumerable<Table> source, string location, List<TableData> result)
    {
        foreach (var table in source)
        {
            // Enumerate physical cells. The engine normalizes horizontal merges and marks
            // vertically merged continuation cells with Length == -1; reading their Text throws.
            var cells = table.Cells.Cast<TableCell>().Select(cell => new TableCellData(cell.Row, cell.Column,
                cell.Length < 0 ? null : cell.Text, cell.Width, cell.Length < 0)).ToArray();
            var rows = table.Rows.Cast<TableRow>().Select(row => new TableRowData(row.Row, row.IsHeader,
                cells.Where(cell => cell.Row == row.Row).ToArray())).ToArray();
            result.Add(new(result.Count + 1, location, table.ID, table.NestedLevel, table.Rows.Count, table.Columns.Count, rows));
        }
    }

}
