open NanoXLSX
open System.IO
open System.Collections.Generic


type Row = { Ean: string; Count: int }

let newRow ean count = { Ean = ean; Count = count }

// let displayRow row =
//     printfn "Address %s Value: %s" row.Ean row.Count

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


let getTableBorder (upperBound: int) (cells: Dictionary<string, Cell>) =
    let addA (n: int) = "A" + string n
    let containsKey n = n |> addA |> cells.ContainsKey
    let getItem key : Cell = cells.Item key

    let isNotHeader n =
        n
        |> addA
        |> getItem
        |> fun (cell: Cell) -> string cell.Value <> "LP."

    let contains =
        [ 1 .. upperBound ]
        |> List.map (fun item -> containsKey item && isNotHeader item)

    List.zip [ 1 .. upperBound ] (contains)
    |> List.filter snd
    |> List.map fst
    |> splitList (fun m n -> m + 1 <> n)
    |> printListOfLists




let cellInfo (cell: Cell) =
    let value = cell.Value.ToString()
    printfn "Address %s Value: %s " cell.CellAddress value

let printAllCells cells =
    Seq.iter (fun (pair: KeyValuePair<string, Cell>) -> cellInfo pair.Value) cells


let readCells (cells: Dictionary<string, Cell>) =
    let ean = cells.Item "B13"
    let count = cells.Item "F13"
    printfn "ean: %s ilosc %s" (string ean.Value) (string count.Value)

let readExcel (filename: string) =
    let wb = NanoXLSX.Workbook.Load(filename)

    // let a =
    getTableBorder (wb.CurrentWorksheet.GetLastRowNumber() + 1) wb.CurrentWorksheet.Cells

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
    // readExcel (Array.head argv)
    // writeExcel (Array.head argv)

    if isNull argv || Array.length argv <> 1 then
        printfn "%s" usage
        1
    else
        convertExcel (Array.head argv) |> readExcel
        0
