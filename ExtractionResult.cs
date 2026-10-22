namespace DocxExtraction;

public sealed record ExtractionResult(
    string SourceFile,
    DocumentMetadata Metadata,
    IReadOnlyList<ContentControlData> ContentControls,
    IReadOnlyList<LegacyFormFieldData> LegacyFormFields,
    IReadOnlyList<ApplicationFieldData> ApplicationFields,
    IReadOnlyList<TableData> Tables,
    IReadOnlyList<SectionData> Sections,
    IReadOnlyList<string> Warnings);

public sealed record DocumentMetadata(string Title, string Author, string Subject, string[] Keywords,
    DateTime? Created, string CreatorApplication, IReadOnlyDictionary<string, object?> CustomProperties);
public sealed record ChoiceData(string DisplayText, string Value);
public sealed record ContentControlData(string Location, long Id, string Title, string Tag, string Type,
    string Text, bool? Checked, DateOnly? Date, IReadOnlyList<ChoiceData> Choices);
public sealed record LegacyFormFieldData(string Location, string Name, string Type, string Text,
    bool? Checked, IReadOnlyList<string> Items);
public sealed record ApplicationFieldData(string Location, string Type, string Name, string Text, string[] Parameters);
public sealed record TableData(int Number, string Location, int Id, int NestedLevel, int RowCount,
    int ColumnCount, IReadOnlyList<TableRowData> Rows);
public sealed record TableRowData(int Number, bool IsHeader, IReadOnlyList<TableCellData> Cells);
public sealed record TableCellData(int Row, int Column, string? Text, int WidthTwips, bool IsMergedContinuation);
public sealed record SectionData(int Number, int Start, int Length, string Text, string BreakKind,
    bool Landscape, IReadOnlyList<HeaderFooterData> HeadersAndFooters);
public sealed record HeaderFooterData(string Type, bool ConnectedToPrevious, string Text);
