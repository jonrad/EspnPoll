open System
open Microsoft.Office.Interop
open System.Runtime.InteropServices
open FSharp.Data

type History = JsonProvider<"""[{"selectionId":1,"teamId":10,"undo":false,"player":{"playerId":13934,"firstName":"Antonio","lastName":"Brown","positionId":3,"proTeamId":23},"slotCategoryId":4,"isKeeper":false},{"selectionId":2,"teamId":2,"undo":false,"player":{"playerId":16733,"firstName":"Odell","lastName":"Beckham Jr.","positionId":3,"proTeamId":19},"slotCategoryId":4,"isKeeper":false}]""">

let worksheetRead (sheet : Excel.Worksheet) cell =
    sheet.Cells.Range(cell).Value2

let readColumn (range : Excel.Range) =
    let all = (range.Value2 :?> obj[,])
    [1..all.Length] |> Seq.map (fun x -> all.[x,1])

let worksheetWrite (sheet : Excel.Worksheet) cell value =
    sheet.Cells.Range(cell).Value2 <- value

let getApp () =
    try
        Marshal.GetActiveObject "Excel.Application" :?> Excel.Application
    with
        | ex -> 
            eprintfn "Couldn't find open excel application. If excel is open, try running this as admin"
            raise ex

let findWorksheet (app : Excel.Application) =
    let workbookMaybe = 
        [0..app.Workbooks.Count] 
        |> Seq.map (fun i -> app.Workbooks.[i+1])
        |> Seq.tryFind (fun i -> i.Name.StartsWith("CSG"))

    let workbook =
        match workbookMaybe with
        | Some w -> w
        | None ->
            failwith "Could not find proper excel spreadsheet. Please make sure you have the CSG spreadsheet open"

    workbook.Worksheets.Item 1 :?> Excel.Worksheet

let rec getToken () =
    printf "Token: "
    let token = Console.ReadLine()
    if (new Text.RegularExpressions.Regex("^\d(:\d+){4}$")).IsMatch(token) then
        let league = token.Split(':').[1]
        (league, token)
    else
        printfn "Invalid Token"
        getToken ()

[<EntryPoint>]
let main argv = 
    printfn "Please view directions at https://github.com/jonrad/EspnPoll"

    let firstRow = 12
    let espnColumn = "C"
    let pickColumn = "J"

    let sheet = findWorksheet(getApp())

    let lastRow =
        readColumn (sheet.Cells.Range("A:A"))
        |> Seq.mapi (fun i x -> (i + 1, x))
        |> Seq.takeWhile (fun (i, x) -> i < firstRow || x <> null)
        |> Seq.last
        |> fst

    let rowMap = 
        readColumn (sheet.Cells.Range("C:C")) 
        |> Seq.take lastRow
        |> Seq.mapi (fun i x -> (i + 1, x))
        |> Seq.filter (fun (i, _) -> i >= firstRow)
        |> Seq.filter (fun (_, x) -> x <> null)
        |> Seq.map (fun (i, x) -> (i, x.ToString()))
        |> Seq.filter (fun (_, x) -> x <> null && x <> "" && x <> "0")
        |> Seq.map (fun (i, x) -> (x, i))
        |> Map.ofSeq

    let read = worksheetRead sheet
    let write = worksheetWrite sheet

    let processHistory (history : History.Root) =
        let pickNumber = history.SelectionId
        let name = sprintf "%s %s" history.Player.FirstName history.Player.LastName
        printfn "Pick #%d: %s" pickNumber name

        let row = rowMap.TryFind name
        match row with
        | Some i -> write (sprintf "%s%d" pickColumn i) pickNumber
        | None -> printfn "Could not find %s in the spreadsheet" name
        ()

    let processText text =
        // Thoughts: Would be fun to handle this using something like 
        // https://github.com/sebastienros/jint
        // but that seems heavy handed for now
        let regex = new Text.RegularExpressions.Regex """"pickHistory":(\[[^\]]+\])"""
        let matches = regex.Matches text
        let histories = 
            [0..matches.Count - 1] 
            |> Seq.map (fun x -> matches.[x].Groups.[1].Value) 
            |> Seq.map History.Parse
            |> Seq.concat

        let regexSynch = new Text.RegularExpressions.Regex """"synchTime":(\d+)"""
        let synchMatches = regexSynch.Matches text

        let synchTime = if synchMatches.Count > 0 then Int64.Parse(synchMatches.[synchMatches.Count - 1].Groups.[1].Value) else 0L

        (synchTime, histories)

    let pollWeb league token synchTime =
        let url = sprintf "http://fantasydraft.espn.go.com/league-%s/extdraft/json/POLL?&poll=%d&token=%s&r=522" league synchTime token
        use wc = new Net.WebClient()
        wc.DownloadString(url) |> processText

    let runApp getToken poll (sleepTime : int) =
        let (league, token) = getToken ()

        let rec loop (synchTime : int64) =
            let (newSynchTime, histories) = poll league token synchTime 
            histories |> Seq.iter processHistory

            printfn "polled %d, new synch time %d" synchTime newSynchTime
            Threading.Thread.Sleep sleepTime
            loop (if newSynchTime > synchTime then newSynchTime else synchTime)

        loop 0L

    runApp getToken pollWeb 5000
    0
