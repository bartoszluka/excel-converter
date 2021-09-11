module Strings

open System

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
