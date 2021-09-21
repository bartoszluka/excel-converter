let (|Even|Odd|) n = if n % 2 = 0 then Even n else Odd n

let rec threeNplus1Terminating n =
    seq {
        yield n

        if n <> 1 then

            yield!
                match n with
                | Even n -> threeNplus1Terminating (n / 2)
                | Odd n -> threeNplus1Terminating (3 * n + 1)
    }

let rec threeNplus1Inf n =
    seq {
        yield n

        yield!
            match n with
            | Even n -> threeNplus1Inf (n / 2)
            | Odd n -> threeNplus1Inf (3 * n + 1)
    }

let rec pseudoFib n m =
    seq {
        yield n
        yield! pseudoFib m (n + m)
    }

let fibonacci = pseudoFib 0 1
let fibonacciNth n = fibonacci |> Seq.take n |> Seq.last

let min a b = if a < b then a else b
let max a b = if a > b then a else b

let rec gcd n m =
    let bigger = max n m
    let smaller = min n m

    match smaller with
    | 0 -> bigger
    | 1 -> 1
    | _ -> gcd smaller (bigger - smaller)

let fork binary unary1 unary2 x = binary (unary1 x) (unary2 x)

let numbers = [ 1 .. 100000 ]

let lengths =
    List.map (threeNplus1Terminating >> Seq.length)

let maxElement =
    List.map (threeNplus1Terminating >> Seq.max)

// numbers
// |> fork List.zip id maxElement
// |> List.maxBy snd
// |> printfn "%O"

// let pair x y = (x, y)

// pseudoFib 0 1
// |> Seq.takeWhile ((>) 1_000_000)
// |> fork pair Seq.last Seq.length
// |> printfn "%O"

// gcd 60 12 |> printfn "%i"
// gcd 50 200 |> printfn "%i"

fibonacciNth 9 |> printfn "%i"
