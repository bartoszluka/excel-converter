module ExcelConverter

open NanoXLSX
open System
open System.IO
open ExcelDataReader
open Strings

type Row = { Ean: string; Count: int }

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
    try
        array |> Array.take (Array.length array - n)
    with
    | _ -> [||]

let tryInt str =
    try
        Int32.Parse str |> Some
    with
    | _ -> None

let readExcelKakadu inputFile =
    // 94 ms
    let rawData = readExcelToSeq inputFile

    // 7 ms
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


    // 5 ms
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

    let joinProductsAndCustomers (name, rest) = (name, Array.zip products rest)

    // 49 ms
    customers
    |> Array.map (joinProductsAndCustomers >> toRows)
    |> List.ofArray


let tablesFromExcel =
    readExcelToSeq
    >> takeColumns [ 'B'; 'C'; 'F' ]
    >> Seq.map arrayToTriplet
    >> List.ofSeq
    >> splitInput isRowEmpty
    >> List.map (createTable >> simpleTableOfTable)

let writeExcel directoryName (tables: SimpleTable list) =

    let addRow (wb: Workbook) (row: Row) : unit =
        wb.CurrentWorksheet.AddNextCell(row.Ean)
        wb.CurrentWorksheet.AddNextCell(Cell(row.Count, Cell.CellType.NUMBER))
        wb.CurrentWorksheet.GoToNextRow()

    let directory = Directory.CreateDirectory directoryName

    let createFile table =

        let fileName =
            table.Name
            |> replaceForbiddenWith ' '
            |> normalizeWhiteSpace

        let createdFileName =
            sprintf "%s\\%s.xlsx" (directory.FullName) fileName

        let wb =
            Workbook(
                createdFileName,
                "Arkusz 1",
                WorkbookMetadata = Metadata(Creator = "Katarzyna Stasiak", Application = "Microsoft Excel")
            )

        wb.CurrentWorksheet.AddNextCell("Kod")
        wb.CurrentWorksheet.AddNextCell("Ilosc")
        wb.CurrentWorksheet.AddNextCell("jm")
        wb.CurrentWorksheet.GoToNextRow()

        table.Data |> List.iter (addRow wb)
        wb.Save()
        createdFileName

    tables |> List.map createFile

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



let convertWith toTables (inputFile: string) inputDict =
    try
        // 168 ms
        let tables = toTables inputFile

        // 28 ms
        let dict = createDictionary inputDict

        let directory = Path.GetDirectoryName inputFile

        // 141 ms
        tables
        |> List.map (replaceNames dict)
        |> writeExcel (directory + "/zamienione")
        |> Ok

    // everything lasts 343 ms

    with
    | :? IndexOutOfRangeException -> Error "incorrect file format"
    | :? NanoXLSX.Exceptions.IOException as ex -> Error <| "Could not create file: " + ex.Message
    | :? IOException as ex ->
        Error
        <| "Could not read one or more files: " + ex.Message

let convertKakadu = convertWith readExcelKakadu

let convert = convertWith tablesFromExcel
