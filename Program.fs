module ExcelConverter

open NanoXLSX
open System
open System.IO
open ExcelDataReader

type Row = { Ean: string; Count: int }

let newRow ean count = { Ean = ean; Count = count }

type Table =
    { NumerZamowienia: string
      DataWystawienia: DateTime
      Realizacja: DateTime
      Waluta: string
      WartoscNetto: float
      FirmaNazwa: string
      FirmaGln: string
      PlatnikNazwa: string
      PlatnikGln: string
      KontrahentNazwa: string
      KontrahentEan: string
      Data: Row list }

let rowOfPair (ean, count) = { Ean = ean; Count = int count }

let newTable (arr: string []) (rowData: (string * string) list) =
    { NumerZamowienia = arr.[0]
      DataWystawienia = DateTime.Parse arr.[1]
      Realizacja = DateTime.Parse arr.[2]
      Waluta = arr.[3]
      WartoscNetto = float arr.[4]
      FirmaNazwa = arr.[5]
      FirmaGln = arr.[6]
      PlatnikNazwa = arr.[7]
      PlatnikGln = arr.[8]
      KontrahentNazwa = arr.[9]
      KontrahentEan = arr.[10]
      Data = List.map rowOfPair rowData }

let isRowEmpty (a, b, c) =
    let isEmpty = (=) ""
    isEmpty a && isEmpty b && isEmpty c

let rec splitInTwo pred =
    function
    | [] -> ([], [])
    | x :: xs when pred x -> [], xs
    | x :: xs ->
        let (splitPart, rest) = splitInTwo pred xs
        (x :: splitPart, rest)

let rec splitInput pred input =
    match splitInTwo pred input with
    | xs, [] -> [ xs ]
    | xs, ys -> xs :: (splitInput pred ys)

let createTable rawInput =
    let (list1, list2) =
        splitInTwo ((=) ("EAN", "NAZWA", "ILOSC")) rawInput

    let takeMiddle (_, x, _) = x
    let skipMiddle (x, _, y) = (x, y)

    let metadata =
        list1 |> List.map takeMiddle |> Array.ofList

    let rows = list2 |> List.map skipMiddle

    newTable metadata rows

let readExcelToSeq column1 column2 column3 inputFile =
    // this is necessary for .NET Core to work
    Text.Encoding.RegisterProvider(Text.CodePagesEncodingProvider.Instance)

    // read input file into a data set
    let dataSet =
        ExcelReaderFactory
            .CreateReader(File.Open(inputFile, FileMode.Open, FileAccess.Read))
            .AsDataSet()

    let toIndex char = int (Char.ToUpper char) - int 'A'

    // convert data set into a sequence
    seq {
        for row in dataSet.Tables.[0].Rows do
            let s letter = string row.ItemArray.[toIndex letter]
            yield (s column1, s column2, s column3)
    }

let tablesFromExcel =
    readExcelToSeq 'B' 'C' 'F'
    >> List.ofSeq
    >> splitInput isRowEmpty
    >> List.map createTable

let normalizeWhiteSpace input =
    let toCharList =
        function
        | str when String.IsNullOrEmpty str -> []
        | str -> str.ToCharArray() |> List.ofArray

    let foldr folder list state = List.foldBack folder state list
    let lastChar (str: string) = str.[str.Length - 1]

    let bothWhiteSpace x y =
        Char.IsWhiteSpace x && Char.IsWhiteSpace y

    match toCharList input with
    | [] -> ""
    | first :: rest ->
        string first
        + (first :: rest
           |> List.pairwise
           |> List.fold
               (fun acc (x, y) ->
                   if bothWhiteSpace x y then
                       acc
                   else
                       acc + (string y))
               "")

let replaceForbiddenWith c =
    let forbiddenChars =
        [ '<'
          '>'
          ':'
          '\"'
          '/'
          '\\'
          '|'
          '?'
          '*' ]

    let inline (?<) item list = List.contains item list

    let (|Forbidden|Safe|) char =
        if char ?< forbiddenChars then
            Forbidden char
        else
            Safe char

    String.map
        (function
        | Forbidden _ -> c
        | Safe char -> char)

let writeExcel directoryName (tables: Table list) =

    let addRow (wb: Workbook) (row: Row) : unit =
        wb.CurrentWorksheet.AddNextCell(row.Ean)
        wb.CurrentWorksheet.AddNextCell(Cell(row.Count, Cell.CellType.NUMBER))
        wb.CurrentWorksheet.GoToNextRow()

    let directory = Directory.CreateDirectory directoryName

    let createFile index table =

        let fileName =
            table.KontrahentNazwa
            |> replaceForbiddenWith ' '
            |> normalizeWhiteSpace


        // let createdFileName =
        //     sprintf "%s\\%s_%d.xlsx" (directory.FullName) fileName (index + 1)
        let createdFileName =
            sprintf "%s\\%s.xlsx" (directory.FullName) fileName

        let wb =
            Workbook(createdFileName, "Arkusz 1", WorkbookMetadata = Metadata(Application = "Microsoft Excel"))

        wb.CurrentWorksheet.AddNextCell("KOD")
        wb.CurrentWorksheet.AddNextCell("ILOŚĆ")
        wb.CurrentWorksheet.AddNextCell("jm")
        wb.CurrentWorksheet.GoToNextRow()

        table.Data |> List.iter (addRow wb)
        wb.Save()
        createdFileName

    tables |> List.mapi createFile

let createDictionary =
    let doesNotMatter = 'D'
    let skipLast (x, y, _) = (x, y)

    readExcelToSeq 'C' 'D' doesNotMatter
    >> Seq.map skipLast
    >> Seq.skip 2
    >> Map.ofSeq

let replaceNames map table =
    let updateRow row maybeEan =
        match maybeEan with
        | None -> row
        | Some ean -> { row with Ean = ean }

    let updateTable table newData = { table with Data = newData }

    table.Data
    |> List.map (fun row -> Map.tryFind row.Ean map |> updateRow row)
    |> updateTable table

let convert inputFile inputDict =
    try
        let dict = createDictionary inputDict
        let tables = tablesFromExcel inputFile
        let directory = Path.GetDirectoryName inputFile

        tables
        |> List.map (replaceNames dict)
        |> writeExcel (directory + "/zamienione")
        |> Ok
    with
    | :? IndexOutOfRangeException -> Error "incorrect file format"
    | :? NanoXLSX.Exceptions.IOException as ex -> Error <| "Could not create file: " + ex.Message
    | :? IOException as ex ->
        Error
        <| "Could not read one or more files: " + ex.Message
