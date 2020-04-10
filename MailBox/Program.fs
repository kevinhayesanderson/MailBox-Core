open Menu
open MailUtils

[<EntryPoint>]
let main _argv =
    match mainMenu() with
        | "1" -> 
            let sendMailAns = List.ofSeq(retrieveAnswers sendMailQuestions)
            sendMail sendMailAns.[0] sendMailAns.[1] sendMailAns.[2] sendMailAns.[3] sendMailAns.[4]
        | "2" ->
            let serverType = match incomingServerTypeMenu() with
                                | "1" ->  IMAP
                                | "2" ->  POP3
                                | _ -> 
                                    logInvalidOption()
                                    None
            let receiveMailAns = List.ofSeq(retrieveAnswers receiveMailQuestions)
            receiveMail serverType receiveMailAns.[0] receiveMailAns.[1] receiveMailAns.[2]
        | _ -> 
            logInvalidOption()
    let _key = System.Console.ReadKey()
    0 
