open System
open MailUtils
open Hopac
open Logary
open Logary.Message
open Logary.Configuration
open Logary.Targets

[<EntryPoint>]
let main _argv =

    let menuOptions = [ "Send Mail"; "Read Mail" ]

    let sendMailQuestions =
        [ "Enter valid recipient mailId's(Comma or semicolon Seperated, if multiple):";
        "Enter valid from MailId:";
        "Enter mail subject:";
        "Enter mail body:" ]

    let receiveMailQuestions = [ "Enter Server Name:"; "Enter Username:"; "Enter Password" ]

    let printOptions options = options |> List.iteri (fun i x -> Console.WriteLine("{0}. {1}", (i + 1), x))

    let getOption n =
        Console.Write("Enter option:")
        List.iter (fun (i: int) ->
            Console.Write((if i = n then "{0}" else "{0},"), i)) [ 1 .. n ]
        Console.WriteLine()
        Console.ReadLine().Trim()

    let mainMenu() =
        Console.WriteLine("Menu:")
        menuOptions |> printOptions
        menuOptions.Length |> getOption
    
    let getUserInput (message: string) =
        Console.WriteLine(message)
        Console.ReadLine().Trim()

    let retrieveAnswers questions =
        seq {
            for question in questions do
                yield getUserInput question
        }

    let logary =
        Config.create "Mailbox-Core" "local"
        |> Config.target (LiterateConsole.create LiterateConsole.empty "console")
        |> Config.ilogger (ILogger.Console Debug)
        |> Config.build
        |> run

    let logger = logary.getLogger "Program"

    match mainMenu() with
        | "1" -> 
            let sendMailAns = List.ofSeq(retrieveAnswers sendMailQuestions)
            sendMail sendMailAns.[0] sendMailAns.[1] sendMailAns.[2] sendMailAns.[3]
        | "2" ->
            let receiveMailAns = List.ofSeq(retrieveAnswers receiveMailQuestions)
            receiveMail receiveMailAns.[0] receiveMailAns.[1] receiveMailAns.[2]
        | _ -> 
            event Error "Invalid option"
            |> Logger.logSimple logger
    
    let _key = System.Console.ReadKey()
    0 
