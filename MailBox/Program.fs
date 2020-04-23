open Menu
open MailUtils
open Hopac
open Logary
open Logary.Message
open Logary.Configuration
open Logary.Targets

[<EntryPoint>]
let main _argv =
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
    // sendMailOutlook "" "" "kevin.hayes@ambigai.net"
    
    let _key = System.Console.ReadKey()
    0 
