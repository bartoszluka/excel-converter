module ExcelConverter

open NanoXLSX
open System
open System.IO
open ExcelDataReader

type Row = { Ean: string; Count: int }

let newRow ean count = { Ean = ean; Count = count }

let displayRow row =
    printfn "Ean %s Count: %d" row.Ean row.Count

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

let splitInTwo predicate input =
    let rec splitInTwoHelper pred building rest =
        match rest with
        | [] -> (building, [])
        | x :: xs when pred x -> List.rev building, xs
        | x :: xs -> splitInTwoHelper pred (x :: building) xs

    splitInTwoHelper predicate [] input

let rec splitInput pred input =
    match splitInTwo pred input with
    | xs, [] -> [ List.rev xs ]
    | xs, ys -> xs :: (splitInput pred ys)

let createTable rawInput =
    let (list1, list2) =
        splitInTwo ((=) ("EAN", "NAZWA", "ILOSC")) rawInput

    let metadata =
        list1
        |> List.map (fun (_, value, _) -> value)
        |> Array.ofList

    let rows =
        list2 |> List.map (fun (x, _, y) -> (x, y))

    newTable metadata rows

let readExcelToSeq inputFile =
    // this is necessary for .NET Core to work
    Text.Encoding.RegisterProvider(Text.CodePagesEncodingProvider.Instance)

    // read input file into a data set
    let dataSet =
        ExcelReaderFactory
            .CreateReader(File.Open(inputFile, FileMode.Open, FileAccess.Read))
            .AsDataSet()

    // convert data set into a sequence
    seq {
        for row in dataSet.Tables.[0].Rows do
            let s index = string row.ItemArray.[index]
            // this are columns B, C and F
            yield (s 1, s 2, s 5)
    }

let tablesFromExcel =
    readExcelToSeq
    >> List.ofSeq
    >> splitInput isRowEmpty
    >> List.map createTable

let writeExcel directoryName (tables: Table list) =

    let addRow (wb: Workbook) (row: Row) : unit =
        wb.CurrentWorksheet.AddNextCell(row.Ean)
        wb.CurrentWorksheet.AddNextCell(Cell(row.Count, Cell.CellType.NUMBER))
        wb.CurrentWorksheet.GoToNextRow()

    let directory = Directory.CreateDirectory directoryName

    let createFile index table =
        let fileName =
            table.KontrahentNazwa
            |> String.map
                (function
                | '/' -> ' '
                | '\\' -> ' '
                | notSpace -> notSpace)

        // let createdFileName =
        //     sprintf "%s/%s_%d.xlsx" (directory.FullName) fileName (index + 1)
        let createdFileName =
            sprintf "%s/%s.xlsx" (directory.FullName) fileName

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

let createDictionary (filename: string) =
    Text.Encoding.RegisterProvider(Text.CodePagesEncodingProvider.Instance)

    let reader =
        File.Open(filename, FileMode.Open, FileAccess.Read)
        |> ExcelReaderFactory.CreateReader

    let result = reader.AsDataSet()

    let actualData =
        seq<string * string> {
            for row in result.Tables.[0].Rows do
                let itemArray = row.ItemArray
                yield (string itemArray.[2], string itemArray.[3])
        }

    actualData |> Seq.skip 2 |> Map.ofSeq

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
    let dict = createDictionary inputDict
    let tables = tablesFromExcel inputFile

    let directory = Path.GetDirectoryName inputFile

    tables
    |> List.map (replaceNames dict)
    |> writeExcel (directory + "/zamienione")
    |> Ok
