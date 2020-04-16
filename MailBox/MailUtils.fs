module MailUtils

open System
open System.Text.RegularExpressions
open System.Globalization
open System.Net.Mail
open DnsClient
open System.Linq
open Menu
open MailKit.Net.Imap
open MailKit.Net.Pop3
open MailKit
open System.ComponentModel
open System.Security.Principal

type ServerName =
    | IMAP
    | POP3
    | SMPT
    | None

type ServerType =
    | Incoming
    | Outgoing

type AuthenticationType =
    | AUTH
    | StartTLS
    | SSL

type DefaultPortInfo =
    { Name: ServerName
      Type: ServerType
      Authentication: AuthenticationType
      Port: int }

let defaultServerInfo =
    ([ { Name = SMPT
         Type = Outgoing
         Authentication = AUTH
         Port = 25 }
       { Name = SMPT
         Type = Outgoing
         Authentication = AUTH
         Port = 26 }
       { Name = SMPT
         Type = Outgoing
         Authentication = StartTLS
         Port = 587 }
       { Name = SMPT
         Type = Outgoing
         Authentication = SSL
         Port = 465 }
       { Name = POP3
         Type = Incoming
         Authentication = AUTH
         Port = 110 }
       { Name = POP3
         Type = Incoming
         Authentication = SSL
         Port = 995 }
       { Name = IMAP
         Type = Incoming
         Authentication = AUTH
         Port = 143 }
       { Name = IMAP
         Type = Incoming
         Authentication = SSL
         Port = 993 } ])

let getPort (name: ServerName) (auth: AuthenticationType) =
    defaultServerInfo.FirstOrDefault(fun si -> si.Name.Equals(name) && (si.Authentication.Equals(auth))).Port

let mutable invalidmailIds = List<string>.Empty

let isValidEmail mailId =
    if String.IsNullOrWhiteSpace(mailId) then
        false
    else
        try
            let domainMapper (mt: Match) =
                let idn = IdnMapping()
                let domainName = idn.GetAscii(mt.Groups.[2].Value)
                mt.Groups.[1].Value + domainName

            let mailId =
                Regex.Replace
                    (mailId.Trim(), @"(@)(.+)$", domainMapper, RegexOptions.None, TimeSpan.FromMilliseconds(200.0))
            Regex.IsMatch
                (mailId,
                 @"^(?("")("".+?(?<!\\)""@)|(([0-9a-z]((\.(?!\.))|[-!#\$%&'\*\+/=\?\^`\{\}\|~\w])*)(?<=[0-9a-z])@))"
                 + @"(?(\[)(\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-z][-0-9a-z]*[0-9a-z]*\.)+[a-z0-9][\-a-z0-9]{0,22}[a-z0-9]))$",
                 RegexOptions.IgnoreCase, TimeSpan.FromMilliseconds(250.0))
        with
        | :? RegexMatchTimeoutException -> false
        | :? ArgumentException -> false

let validateMailIds mailIds = List.partition isValidEmail mailIds

let isValidMailAddress (mailAddress: string) =
    match mailAddress.Contains(',') || mailAddress.Contains(';') with
    | true ->
        let invalidMailIds = snd (validateMailIds (List.ofArray (mailAddress.Split(',', ';'))))
        match (invalidMailIds).Any() with
        | true ->
            invalidmailIds <- invalidmailIds @ invalidMailIds
            false
        | _ -> true
    | _ ->
        match isValidEmail mailAddress with
        | true -> true
        | _ ->
            invalidmailIds <- mailAddress :: invalidmailIds
            false

let getCurrentUserName = WindowsIdentity.GetCurrent().Name

let getDomainName (mailId: string) = (mailId |> MailAddress).Host

let getMXRecord domainName = (LookupClient()).Query(domainName, QueryType.MX).Answers.FirstOrDefault()

let getSMPTServerName (record: Protocol.DnsResourceRecord) =
    match record with
    | :? Protocol.MxRecord as mxRecord -> mxRecord.Exchange.Original.TrimEnd('.')
    | _ -> ""

let setUpMailMessage (toAddress: string) fromAddress subject body =
    let mail = new MailMessage()
    mail.From <- MailAddress(fromAddress)
    mail.To.Add(toAddress)
    mail.Subject <- subject
    mail.SubjectEncoding <- Text.Encoding.UTF8
    mail.Body <- body
    mail.BodyEncoding <- Text.Encoding.UTF8
    mail

let setUpSmtpClient host port =
    let smtpClient = new SmtpClient(host, port)
    //smtpClient.Credentials <- new System.Net.NetworkCredential("username", "password")
    smtpClient.UseDefaultCredentials <- true
    smtpClient.EnableSsl <- true
    smtpClient

let printMails client =
    match box client with
    | :? ImapClient as imapClient ->
        for summary in imapClient.Inbox.Fetch(1, imapClient.Inbox.Count - 1, MessageSummaryItems.Full) do
            Console.WriteLine("[summary] {0:D2}: {1}", summary.Index, summary.Envelope.Subject)
        Console.WriteLine()
        Console.WriteLine("Printed all {0} mails", imapClient.Inbox.Count)
    | :? Pop3Client as pop3Client ->
        for i = 1 to pop3Client.Count do
            let message = pop3Client.GetMessage(i)
            Console.WriteLine("Subject: {0}", message.Subject)
        Console.WriteLine("Printed all {0} mails", pop3Client.Count)
    | _ -> Console.WriteLine("No Mail Found")

let imapPrintMailsCallback =
    (fun (imapClient: ImapClient) ->
        let inbox = imapClient.Inbox
        inbox.Open(FolderAccess.ReadOnly) |> ignore
        Console.WriteLine("Total messages: {0}", inbox.Count)
        printMails imapClient
        imapClient.Disconnect(true))

let setUpIMAPClient serverName (userName: string) password printMailsCallback =
    let imapClient = new ImapClient()
    imapClient.Timeout <- 9000
    imapClient.Connected
    |> Event.add
        (fun args ->
            Console.WriteLine
                ("IMAP Client Connected: Host-{0} Port-{1} SecureSocketOption-{2}", args.Host, args.Port, args.Options))
    imapClient.ConnectAsync(serverName,
                            IMAP
                            |> getPort
                            <| SSL, true).GetAwaiter().GetResult()
    imapClient.Authenticated
    |> Event.add (fun args ->
        Console.WriteLine(args.Message)
        printMailsCallback imapClient)
    imapClient.AuthenticateAsync(userName, password).GetAwaiter().GetResult()

let pop3PrintMailsCallback =
    (fun (pop3Client: Pop3Client) ->
        Console.WriteLine("Total messages: {0}", pop3Client.Count)
        printMails pop3Client
        pop3Client.Disconnect(true))

let setUpPOP3Client serverName (userName: string) password printMailsCallback =
    let pop3Client = new Pop3Client()
    pop3Client.Timeout <- 9000
    pop3Client.Connected
    |> Event.add
        (fun args ->
            Console.WriteLine
                ("POP3 Client Connected: Host-{0} Port-{1} SecureSocketOption-{2}", args.Host, args.Port, args.Options))
    pop3Client.ConnectAsync(serverName,
                            POP3
                            |> getPort
                            <| SSL, true).GetAwaiter().GetResult()
    pop3Client.Authenticated
    |> Event.add (fun args ->
        Console.WriteLine(args.Message)
        printMailsCallback pop3Client)
    pop3Client.AuthenticateAsync(userName, password).GetAwaiter().GetResult()

let sendMailCallback (args: AsyncCompletedEventArgs) =
    let token = args.UserState
    if args.Cancelled
    then Console.WriteLine("[{0}] Send canceled.", token)
    elif not (isNull args.Error)
    then Console.WriteLine("[{0}] {1}", token, args.Error.ToString())
    else Console.WriteLine("Message sent.")

let sendMail toAddress fromAddress subject body =
    try
        match (isValidMailAddress toAddress) && (isValidMailAddress fromAddress) with
        | true ->
            let mailMessage = setUpMailMessage toAddress fromAddress subject body

            let smptClient =
                setUpSmtpClient
                    (fromAddress
                     |> getDomainName
                     |> getMXRecord
                     |> getSMPTServerName)
                    (SMPT
                     |> getPort
                     <| SSL)
            smptClient.SendCompleted.AddHandler(fun _ e -> sendMailCallback e)
            smptClient.SendMailAsync(mailMessage).GetAwaiter().GetResult()
        | false ->
            logError "invalid Mailid's"
            List.iter (logFunc "Error") invalidmailIds
    with ex -> logError ex.Message

let getServerType (serverName: string) =
    match serverName with
    | serverName when serverName.Contains("IMAP") -> IMAP
    | serverName when serverName.Contains("imap") -> IMAP
    | serverName when serverName.Contains("POP") -> POP3
    | serverName when serverName.Contains("pop") -> POP3
    | _ -> None

let receiveMail serverName userName password =
    try
        match getServerType (serverName) with
        | IMAP -> setUpIMAPClient serverName userName password imapPrintMailsCallback
        | POP3 -> setUpPOP3Client serverName userName password pop3PrintMailsCallback
        | _ -> logError "Invalid servertype"
    with ex -> logError ex.Message
