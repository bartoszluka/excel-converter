module ExcelConverter

open NanoXLSX
open System.IO
open System.Collections.Generic

type Row = { Ean: string; Count: int }

let newRow ean count = { Ean = ean; Count = count }

let displayRow row =
    printfn "Ean %s Count: %d" row.Ean row.Count

type Table = Rows of Row list

let displayTable (Rows rows) = rows |> (List.iter displayRow)

let displayTables tables = tables |> (List.iter displayTable)

let unPair f (x, y) = f x y

let splitWhen (f: 'a -> 'a -> bool) (list: list<'a>) =
    if (List.isEmpty list) then
        ([], [])

    else

        let firstIndex =
            list
            |> List.pairwise
            |> List.tryFindIndex (unPair f)

        match firstIndex with
        | None -> (list, [])
        | Some n -> List.splitAt (if n = 0 then 0 else n + 1) list

let rec splitList f list =
    match splitWhen f list with
    | ([], []) -> []
    | ([], _) -> []
    | (a, []) -> [ a ]
    | (a, b) -> a :: splitList f b

let createTable =
    List.map (fun (ean, count) -> newRow ean count)

let toRecords dict (list: int list list) =
    let letters = [ 'A' .. 'K' ] |> List.map string


    let getLetterForKey key header letter =
        (Map.tryFind (letter + string header) dict)
        |> Option.bind
            (fun item ->
                if string item = key then
                    Some letter
                else
                    None)

    let firstHeader = list |> List.head |> List.head

    let findLetter key =
        List.map (getLetterForKey key firstHeader) letters
        |> List.tryFind Option.isSome
        |> Option.flatten

    let maybeEan = findLetter "EAN" //B
    let maybeCount = findLetter "ILOSC" //F

    let getValues eanHeader countHeader (rowNumbers: int list) =
        rowNumbers
        |> List.map
            (fun item -> ((Map.find (eanHeader + string item) dict), int (Map.find (countHeader + string item) dict)))

    let noHeaders = list |> List.map List.tail

    match maybeEan, maybeCount with
    | Some ean, Some count ->
        noHeaders
        |> List.map ((getValues ean count) >> createTable >> Rows)
        |> Some
    | _ -> None


let flip f x y = f y x

let getTableBorders (upperBound: int) (cells: Map<string, string>) =
    let addA (n: int) = "A" + string n

    let containsKey n = n |> addA |> flip Map.containsKey cells

    let isNotHeader n =
        n
        |> addA
        |> flip Map.find cells
        |> fun cell -> cell <> "LP."

    let contains =
        [ 1 .. upperBound ]
        // |> List.map (fun item -> containsKey item && isNotHeader item)
        |> List.map containsKey

    List.zip [ 1 .. upperBound ] (contains)
    |> List.filter snd
    |> List.map fst
    |> splitList (fun m n -> m + 1 <> n)
    |> toRecords cells



let cellInfo (cell: Cell) =
    let value = cell.Value.ToString()
    printfn "Address %s Value: %s " cell.CellAddress value

let printAllCells cells =
    Seq.iter (fun (pair: KeyValuePair<string, Cell>) -> cellInfo pair.Value) cells


let readCells (cells: Dictionary<string, Cell>) =
    let ean = cells.Item "B13"
    let count = cells.Item "F13"
    printfn "ean: %s ilosc %s" (string ean.Value) (string count.Value)

let toMap dictionary =
    (dictionary :> seq<_>)
    |> Seq.map (|KeyValue|)
    |> Map.ofSeq

let excelToTables (filename: string) =
    let wb = NanoXLSX.Workbook.Load(filename)

    let map =
        toMap wb.CurrentWorksheet.Cells
        |> Map.map (fun _ (value: Cell) -> string value.Value)

    getTableBorders (wb.CurrentWorksheet.GetLastRowNumber() + 1) map

let excelOldToNew filename =
    let output = Path.ChangeExtension(filename, "xlsx")
    use workbook = new Spire.Xls.Workbook()
    workbook.LoadFromFile filename
    workbook.SaveToFile(output, Spire.Xls.ExcelVersion.Version2016)
    output

let writeExcel (filename: string) directoryName (tables: Table list) =
    let addRow (wb: Workbook) (row: Row) : unit =
        wb.CurrentWorksheet.AddNextCell(row.Ean)
        wb.CurrentWorksheet.AddNextCell(row.Count)
        wb.CurrentWorksheet.GoToNextRow()

    let directory = Directory.CreateDirectory directoryName

    let createFile index (Rows rows) =
        let createdFileName =
            sprintf "%s/%s.xlsx" (directory.FullName) (filename + (string (1 + index)))

        let wb = Workbook(createdFileName, "Arkusz 1")

        wb.CurrentWorksheet.AddNextCell("KOD")
        wb.CurrentWorksheet.AddNextCell("ILOŚĆ")
        wb.CurrentWorksheet.AddNextCell("jm")
        wb.CurrentWorksheet.GoToNextRow()

        rows |> List.iter (addRow wb)
        wb.Save()
        createdFileName

    tables |> List.mapi createFile


// let removeKeys keys map =
//     Map.filter (fun key _ -> not <| List.contains key keys) map

// let convertExcel2 filename =
//     let output = filename + "_copy"
//     use workbook = new Spire.Xls.Workbook()
//     workbook.LoadFromFile filename
//     workbook.SaveToFile(output, Spire.Xls.ExcelVersion.Version2016)
//     output

// let readDictionary (filename: string) =
//     let workSheet =
//         NanoXLSX.Workbook.Load(filename).CurrentWorksheet

//     let cells = workSheet.Cells
//     let upperBound = workSheet.GetLastRowNumber() + 1

//     let changeCellType (cell: Cell) =
//         // cell.DataType <- Cell.CellType.STRING
//         string cell.Value

//     let map =
//         cells
//         |> toMap
//         |> Map.map (fun _ (value: Cell) -> changeCellType value)

//     let headers =
//         List.allPairs [ "A"; "B"; "C"; "D"; "E"; "F" ] [ 1 .. 2 ]
//         |> List.map (fun pair -> fst pair + string (snd pair))

//     let firstColumns =
//         List.allPairs [ "A"; "B" ] [ 3 .. upperBound ]
//         |> List.map (fun pair -> fst pair + string (snd pair))

//     cells.Values
//       |> removeKeys headers
//       |> removeKeys firstColumns

// let printMap map =
//     map
//     |> Map.iter (fun key value -> printfn "(%s: %s)" key value)

// let printSeq seq =
//     seq |> Seq.iter (fun value -> printfn "%s" value)

// let printCells cells =
//     cells
//     |> Seq.iter (fun (value: Cell) -> printfn "%s" (string value.Value))

let createDictionary (filename: string) =
    // try
    let lineToPair (line: string) =
        line.Split '\t'
        |> fun array -> (array.[0], array.[1])

    File.ReadAllLines filename
    |> Array.tail
    |> Array.map lineToPair
    |> Map.ofArray
//     |> Ok
// with
// | :? System.IndexOutOfRangeException -> Error "wrong format for the dictionary"
// | :? IOException -> "cannot read file " + filename |> Error
// | _ -> Error "some other error idk"



let replaceNames map (Rows rows) =
    let updateRow row maybeEan =
        match maybeEan with
        | None -> row
        | Some ean -> { row with Ean = ean }

    rows
    |> List.map (fun row -> Map.tryFind row.Ean map |> updateRow row)
    |> Rows

let convert inputFile inputDict =
    let dict = createDictionary inputDict
    let converted = inputFile |> excelOldToNew
    let tables = converted |> excelToTables

    let directory = Path.GetDirectoryName inputFile

    let outFiles =
        match tables with
        | None -> Error "wrong format"
        | Some t ->
            t
            |> List.map (replaceNames dict)
            |> writeExcel "zmienione" (directory + "/zamienione")
            |> Ok

    File.Delete converted |> ignore
    outFiles
