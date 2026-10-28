// Power Query (M) for the "SourceData" query.
// Reads the path from the SourcePath named cell, so the source can be moved
// or renamed by editing one cell. Works with .xlsx and .csv sources.
// A file name alone is resolved against the folder of this workbook (WorkbookFolder named cell).
let
    RawPath = Text.Trim(Text.From(Excel.CurrentWorkbook(){[Name = "SourcePath"]}[Content]{0}[Column1])),
    Folder = Text.From(Excel.CurrentWorkbook(){[Name = "WorkbookFolder"]}[Content]{0}[Column1]),
    IsAbsolute = Text.Contains(RawPath, ":\") or Text.StartsWith(RawPath, "\\"),
    FullPath = if IsAbsolute then RawPath else Folder & RawPath,
    IsCsv = Text.EndsWith(Text.Lower(FullPath), ".csv"),
    Data =
        if IsCsv then
            Table.PromoteHeaders(
                Csv.Document(File.Contents(FullPath), [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
                [PromoteAllScalars = true]
            )
        else
            Table.SelectRows(Excel.Workbook(File.Contents(FullPath), true), each [Kind] = "Sheet"){0}[Data],
    // Drop rows that are completely empty (formatted but unused rows in the source)
    IsBlank = (v) => v = null or (v is text and Text.Trim(v) = ""),
    NoBlankRows = Table.SelectRows(Data, each not List.AllTrue(List.Transform(Record.FieldValues(_), IsBlank))),
    // Drop columns that are completely empty (e.g. "Column7" from formatted but unused cells)
    KeepCols = List.Select(
        Table.ColumnNames(NoBlankRows),
        each not List.AllTrue(List.Transform(Table.Column(NoBlankRows, _), IsBlank))
    ),
    Clean = Table.SelectColumns(NoBlankRows, KeepCols),
    // Show a "Date" column as a date when the source has one
    Typed = if List.Contains(Table.ColumnNames(Clean), "Date")
        then Table.TransformColumns(Clean, {{"Date", each try Date.From(_) otherwise null, type date}})
        else Clean
in
    Typed
