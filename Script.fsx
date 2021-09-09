open System

let rec mapReduce map reduce =
    function
    | [ x ] -> map x
    | x :: xs -> reduce (map x) (mapReduce map reduce xs)

let toCharList =
    function
    | str when String.IsNullOrEmpty str -> []
    | str -> str.ToCharArray() |> List.ofArray


let bothWhiteSpace x y =
    Char.IsWhiteSpace x && Char.IsWhiteSpace y

mapReduce
    (fun (x, y) ->
        if bothWhiteSpace x y then
            ""
        else
            string y)
    (+)
    (toCharList "abcd     efgh" |> List.pairwise)
|> printfn "%s"
