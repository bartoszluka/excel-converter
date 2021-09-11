open System

let rec mapReduce map reduce =
    function
    | [] -> failwith "cannot use on empty array"
    | [ x ] -> map x
    | x :: xs -> reduce (map x) (mapReduce map reduce xs)

let rec tryMapReduce map reduce =
    function
    | [] -> None
    | [ x ] -> Some <| map x
    | x :: xs ->
        match tryMapReduce map reduce xs with
        | Some value -> Some <| reduce (map x) value
        | None -> None

let allWhiteSpace =
    tryMapReduce Char.IsWhiteSpace (&&)
    >> Option.defaultValue true

let rec removeExcessiveSpaces chars =
    let isSpace = Char.IsWhiteSpace

    match chars with
    | [] -> []
    | [ x ] -> if isSpace x then [] else [ x ]
    | x :: y :: rest ->
        match isSpace x, isSpace y with
        | true, true -> removeExcessiveSpaces (y :: rest)
        | _ -> x :: removeExcessiveSpaces (y :: rest)

let rec trimWhiteSpaceFront =
    function
    | [] -> []
    | x :: xs ->
        if Char.IsWhiteSpace x then
            trimWhiteSpaceFront xs
        else
            x :: xs

let toCharList =
    function
    | str when String.IsNullOrEmpty str -> []
    | str -> str.ToCharArray() |> List.ofArray

let toString =
    tryMapReduce string (+) >> Option.defaultValue ""

let normalizeWhiteSpace =
    toCharList
    >> trimWhiteSpaceFront
    >> removeExcessiveSpaces
    >> toString

printfn "%b" (normalizeWhiteSpace "a               b" = "a b")
printfn "%b" (normalizeWhiteSpace "               b" = "b")
printfn "%b" (normalizeWhiteSpace "a    c   d     b" = "a c d b")
printfn "%b" (normalizeWhiteSpace "a               " = "a")
printfn "%b" (normalizeWhiteSpace "   a               " = "a")
// "a b"
// |> toCharList
// |> trimWhiteSpaceFront
// |> toString
// |> (=) "a b"
// |> printfn "%b"
