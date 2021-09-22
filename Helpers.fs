namespace System

module List =

    let rec foldr folding state =
        function
        | [] -> state
        | x :: xs -> folding x (foldr folding state xs)

    let rec tryMapReduce map reduce =
        function
        | [] -> None
        | [ x ] -> Some <| map x
        | x :: xs -> Option.map (reduce (map x)) (tryMapReduce map reduce xs)

    let rec mapReduce map reduce =
        function
        | [] -> failwith "cannot reduce an empty list"
        | [ x ] -> map x
        | x :: xs -> reduce (map x) (mapReduce map reduce xs)

    let rec intersperse sep =
        foldr
            (fun item ->
                function
                | [] -> [ item ]
                | other -> item :: sep :: (intersperse sep other))
            []

// let intersperse sep ls =
//     List.foldBack
//         (fun x ->
//             function
//             | [] -> [ x ]
//             | xs -> x :: sep :: xs)
//         ls
//         []

module Tuple =
    let pair x y = (x, y)
    let applyPair f (x, y) = f x y

module Utils =
    let fork binary unary1 unary2 x = binary (unary1 x) (unary2 x)
    let flip f x y = f y x

module String =
    let allWhiteSpace =
        List.tryMapReduce Char.IsWhiteSpace (&&)
        >> Option.defaultValue true

    let rec removeExcessiveSpacesList chars =
        let isSpace = Char.IsWhiteSpace

        match chars with
        | [] -> []
        | [ x ] -> if isSpace x then [] else [ x ]
        | x :: y :: rest ->
            match isSpace x, isSpace y with
            | true, true -> removeExcessiveSpacesList (y :: rest)
            | _ -> x :: removeExcessiveSpacesList (y :: rest)

    let rec trimWhiteSpaceFrontList =
        function
        | [] -> []
        | x :: xs ->
            if Char.IsWhiteSpace x then
                trimWhiteSpaceFrontList xs
            else
                x :: xs

    let toCharList =
        function
        | str when String.IsNullOrEmpty str -> []
        | str -> str.ToCharArray() |> List.ofArray

    let toString =
        List.tryMapReduce string (+)
        >> Option.defaultValue ""

    let normalizeWhiteSpace =
        toCharList
        >> trimWhiteSpaceFrontList
        >> removeExcessiveSpacesList
        >> toString

    let trimWhiteSpaceFront =
        toCharList >> trimWhiteSpaceFrontList >> toString

    let removeExcessiveSpaces =
        toCharList
        >> removeExcessiveSpacesList
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

        let inline (?<?) item list = List.contains item list

        let (|Forbidden|Safe|) char =
            if char ?<? forbiddenChars then
                Forbidden char
            else
                Safe char

        String.map
            (function
            | Forbidden _ -> c
            | Safe char -> char)

    let isSubstring (subString: string) (superString: string) = superString.Contains subString

    let isCharIn (c: char) (superString: string) = superString.Contains c

    let startsWith (str: string) (superString: string) = superString.StartsWith str

    let endsWith (str: string) (superString: string) = superString.EndsWith str
