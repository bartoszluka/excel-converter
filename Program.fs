open NanoXLSX
open System.IO
open System.Collections.Generic


type Row = { Ean: string; Count: int }

let newRow ean count = { Ean = ean; Count = count }

let displayRow row =
    printfn "Ean %s Count: %d" row.Ean row.Count

type Table = Rows of Row list

let displayTable table =
    let (Rows rows) = table
    rows |> (List.iter displayRow)

let displayTables tables = tables |> (List.iter displayTable)


let splitWhen (f: 'a -> 'a -> bool) (list: list<'a>) =
    if (List.isEmpty list) then
        ([], [])

    else
        let helper pair = f (fst pair) (snd pair)

        let firstIndex =
            list |> List.pairwise |> List.tryFindIndex helper

        match (firstIndex) with
        | None -> (list, [])
        | Some n -> List.splitAt (if n = 0 then 0 else n + 1) list
// |> (fun pair -> fst pair :: splitWhen f (snd pair))

let printListOfLists listOfLists =
    List.iter (fun (list: 'a list) -> printfn "%A" list) listOfLists


let rec splitList f list =
    match splitWhen f list with
    | ([], []) -> []
    | ([], _) -> []
    | (a, []) -> [ a ]
    | (a, b) -> a :: splitList f b

let createTable pairsList =
    pairsList
    |> List.map (fun pair -> newRow (fst pair) (snd pair))

let toRecords map (list: int list list) =
    let letters =
        [ "A"
          "B"
          "C"
          "C"
          "D"
          "E"
          "F"
          "G"
          "H"
          "I"
          "J"
          "K" ]


    let getLetterForKey key header letter =
        (Map.tryFind (letter + string header) map)
        |> Option.bind
            (fun item ->
                if string item = key then
                    Some letter
                else
                    None)

    let firstHeader = list |> List.head |> List.head

    // printfn "%A" (getLetterForKey "EAN" 12 "B")
    // (Map.find "B12") |> string |> printfn "%A"
    // Map.iter (printfn "key: %s value %A") map
    let findLetter key =
        List.map (getLetterForKey key firstHeader) letters
        |> List.tryFind Option.isSome
        |> Option.flatten

    let maybeEan = findLetter "EAN" //B
    let maybeCount = findLetter "ILOSC" //F

    let getValues eanHeader countHeader (rowNumbers: int list) =
        rowNumbers
        |> List.map
            (fun item -> ((Map.find (eanHeader + string item) map), int (Map.find (countHeader + string item) map)))

    let noHeaders = list |> List.map List.tail

    match maybeEan, maybeCount with
    | None, _ -> ()
    | _, None -> ()
    | Some ean, Some count ->
        noHeaders
        |> List.map ((getValues ean count) >> createTable >> Rows)
        |> displayTables


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
        |> List.map (fun item -> containsKey item)

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

let readExcel (filename: string) =
    let wb = NanoXLSX.Workbook.Load(filename)

    let map =
        toMap wb.CurrentWorksheet.Cells
        |> Map.map (fun _ (value: Cell) -> string value.Value)

    getTableBorders (wb.CurrentWorksheet.GetLastRowNumber() + 1) map

let convertExcel filename =
    let output = Path.ChangeExtension(filename, "xlsx")
    use workbook = new Spire.Xls.Workbook()
    workbook.LoadFromFile filename
    workbook.SaveToFile(output, Spire.Xls.ExcelVersion.Version2013)
    output

let writeExcel (filename: string) =
    let wb = Workbook(filename, "arkusz 1")
    wb.Save()

let private usage = "expected one .xls file"

[<EntryPoint>]
let main argv =

    if isNull argv || Array.length argv <> 1 then
        printfn "%s" usage
        1
    else
        convertExcel (Array.head argv) |> readExcel
        0
