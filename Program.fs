open NanoXLSX
open System.IO
open System.Collections.Generic

let readExcel (filename: string) =
    let wb = NanoXLSX.Workbook.Load(filename)
    Seq.iter (fun (cell: KeyValuePair<string, Cell>) -> printfn "%s" cell.Key) wb.CurrentWorksheet.Cells

let convertExcel filename =
    let workbook = new Spire.Xls.Workbook()
    workbook.LoadFromFile(filename)
    let output = Path.ChangeExtension(filename, "xlsx")
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
