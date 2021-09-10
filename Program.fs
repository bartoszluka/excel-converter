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

type SimpleTable = { Name: string; Data: Row list }

let simpleTableOfTable table =
    { Name = table.KontrahentNazwa
      Data = table.Data }

let newSimpleTable name rows = { Name = name; Data = rows }

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

let readExcelToSeq inputFile =
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
            yield Array.map string row.ItemArray
    }

let takeColumns columnNames arrays =
    let toIndex char = int (Char.ToUpper char) - int 'A'
    let columns = List.map toIndex columnNames

    arrays
    |> Seq.map (
        Array.indexed
        >> Array.filter (fun (index, _) -> List.contains index columns)
        >> Array.map snd
    )

let arrayToPair (arr: 'a array) = (arr.[0], arr.[1])
let arrayToTriplet (arr: 'a array) = (arr.[0], arr.[1], arr.[2])

let takeEveryNth n array =
    array
    |> Array.indexed
    |> Array.choose
        (fun (index, value) ->
            if index % n = 0 then
                Some value
            else
                None)

let dropLast n array =
    array |> Array.take (Array.length array - n)

let tryInt str =
    try
        Int32.Parse str |> Some
    with
    | _ -> None

let readExcelKakadu inputFile =
    let rawData = readExcelToSeq inputFile

    let products =
        rawData
        |> takeColumns [ 'B' ]
        |> Seq.skip 2
        |> Seq.map (fun arr -> string arr.[0])
        |> Array.ofSeq

    let weirdArrayToTable (arr: string array) =
        let name = Array.head arr
        let rest = Array.skip 2 arr
        (name, rest)

    let customers =
        rawData
        |> Seq.map (Array.skip 2 >> takeEveryNth 2 >> dropLast 2)
        |> Array.transpose
        |> Array.map (weirdArrayToTable)

    let toRows (name, arrays) =
        let toRow (ean, countOrEmpty) =
            match tryInt countOrEmpty with
            | Some int -> Some { Ean = ean; Count = int }
            | None -> None

        let rows = Array.choose toRow arrays

        { Name = name
          Data = List.ofArray rows }

    customers
    |> Array.map (
        (fun (name, rest) -> (name, Array.zip products rest))
        >> toRows
    )
    |> List.ofArray


let tablesFromExcel =
    readExcelToSeq
    >> takeColumns [ 'B'; 'C'; 'F' ]
    >> Seq.map arrayToTriplet
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

let writeExcel directoryName (tables: SimpleTable list) =

    let addRow (wb: Workbook) (row: Row) : unit =
        wb.CurrentWorksheet.AddNextCell(row.Ean)
        wb.CurrentWorksheet.AddNextCell(Cell(row.Count, Cell.CellType.NUMBER))
        wb.CurrentWorksheet.GoToNextRow()

    let directory = Directory.CreateDirectory directoryName

    let createFile index table =

        let fileName =
            table.Name
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
    readExcelToSeq
    >> takeColumns [ 'C'; 'D' ]
    >> Seq.map arrayToPair
    >> Seq.skip 2
    >> Map.ofSeq

let replaceNames map table =
    let updateRow row maybeEan =
        match maybeEan with
        | None -> row
        | Some ean -> { row with Ean = ean }

    let updateTable (table: SimpleTable) newData = { table with Data = newData }

    table.Data
    |> List.map (fun row -> Map.tryFind row.Ean map |> updateRow row)
    |> updateTable table

let convert inputFile inputDict =
    try
        let dict = createDictionary inputDict

        let tables =
            tablesFromExcel inputFile
            |> List.map simpleTableOfTable

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

let convertKakadu inputFile inputDict =
    try
        let dict = createDictionary inputDict

        let tables = readExcelKakadu inputFile

        let directory = Path.GetDirectoryName inputFile

        tables
        |> List.map (replaceNames dict)
        |> writeExcel (directory + "/zamienione2")
        |> Ok
    with
    | :? IndexOutOfRangeException -> Error "incorrect file format"
    | :? NanoXLSX.Exceptions.IOException as ex -> Error <| "Could not create file: " + ex.Message
    | :? IOException as ex ->
        Error
        <| "Could not read one or more files: " + ex.Message
