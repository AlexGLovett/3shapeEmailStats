using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using EAGetMail;
using System.IO;
using System.Text.RegularExpressions;

namespace _3ShapeEmailStatsGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a folder named "inbox" under current directory
            // to save the email retrieved.
            string curpath = @"C:\Nobel\Outlook Inbox Testing";
            string mailbox = String.Format("{0}\\inbox", curpath);
            var emailDict = new Dictionary<string, Dictionary<string, int>>();

            // If the folder is not existed, create it.
            if (!Directory.Exists(mailbox))
            {
                Directory.CreateDirectory(mailbox);
            }

            //MailServer oServer = new MailServer("outlook.office365.com",
            //            "alexander.lovett@nobelbiocare.com", "ALexandier11243!", ServerProtocol.Imap4);

            // for office365 account, please use
            MailServer oServer = new MailServer("outlook.office365.com",
                    "alexander.lovett@nobelbiocare.com", "ALexandier11243!", ServerProtocol.Imap4);
            MailClient oClient = new MailClient("TryIt");

            // If your POP3 server requires SSL connection,
            // Please add the following codes:
            oServer.SSLConnection = true;
            oServer.Port = 993;

            try
            {
                oClient.Connect(oServer);
                Imap4Folder[] folders = oClient.Imap4Folders;
                int count = folders.Length;
                for (int i = 0; i < count; i++)
                {
                    Imap4Folder folder = folders[i];
                    //if (String.Compare("3Shape Prod Inbox", folder.Name, true) == 0)
                    //{
                    //    //select "INBOX" folder
                    //    oClient.SelectFolder(folder);
                    //    break;
                    //}
                    if (String.Compare("INBOX", folder.Name, true) == 0)
                    {
                        //select "INBOX" folder
                        if (folder.SubFolders.Length > 0)
                        {
                            for (int j = 0; j < folder.SubFolders.Length; j++)
                            {
                                if (String.Compare("3Shape Prod Inbox", folder.SubFolders[j].Name, true) == 0)
                                {
                                    oClient.SelectFolder(folder.SubFolders[j]);
                                    MailInfo[] infos = oClient.GetMailInfos();
                                    for (int k = infos.Length - 1; k > 0; k--) //infos.Length - 1
                                    {
                                        MailInfo info = infos[k];
                                        Console.WriteLine("Index: {0}; Size: {1}; UIDL: {2}",
                                            info.Index, info.Size, info.UIDL);

                                        // Receive email from POP3 server
                                        Mail oMail = oClient.GetMail(info);

                                        //Console.WriteLine("From: {0}", oMail.From.ToString());
                                        //Console.WriteLine("Subject: {0}\r\n", oMail.Subject);
                                        //Console.WriteLine("Body: {0}\r\n", oMail.TextBody);

                                        var body = oMail.TextBody;
                                        string[] stringSeparators = new string[] { "\r\n" };
                                        string[] pieces = body.Split(stringSeparators, StringSplitOptions.RemoveEmptyEntries);
                                        var pieceString = String.Join(" | ",pieces);
                                        var _emailFingerPrint = "";
                                        if (pieceString.Contains("Sakagami"))
                                        {
                                            _emailFingerPrint = "SakagamiError";
                                        }
                                        else if (pieceString.Contains("A_Z_Ling"))
                                        {
                                            _emailFingerPrint = "A_Z_LingErrors";
                                        }
                                        else if (pieceString.Contains("DSG_Dahlin_dsmz_12,29,30_W04_kahan_20180501_1154.zip"))
                                        {
                                            _emailFingerPrint = "DSG_DahlinError";
                                        }
                                        else if (pieceString.Contains("Manh_Nguyen.zip"))
                                        {
                                            _emailFingerPrint = "MahnNguyenError";
                                        }
                                        else if (pieceString.Contains("The ERP file") &&
                                            (pieceString.Contains("The process cannot access the file because it is being used by another process")))
                                        {
                                            _emailFingerPrint = "ERP_ERROR-XMLInUseByAnotherProcess";
                                        }
                                        else if (pieceString.Contains("The ERP file"))
                                        {
                                            _emailFingerPrint = "ERP_ERROR-" + String.Join("", pieces.Last().Split(' ').Take(1)).Replace("(", "");
                                        }
                                        else if (pieceString.Contains("Read timed out"))
                                        {
                                            _emailFingerPrint = "ReadTimedOut";
                                        }
                                        else if (pieceString.Contains("Unable to find emergence profile boundary"))
                                        {
                                            _emailFingerPrint = "AnErrorOccurred-Unable to find emergence profile boundary" + String.Join("", pieces[2].Split(' ').Take(3));
                                        }
                                        else if (pieceString.Contains("An error occurred"))
                                        {
                                            _emailFingerPrint = "AnErrorOccurred-" + String.Join("", pieces[2].Split(' ').Take(3));
                                        }
                                 
                                        else
                                        {
                                            _emailFingerPrint = String.Join("",pieces[0].Split(' ').Take(3));
                                        }
                                        var emailFingerPrint = "";
                                        foreach (var print in _emailFingerPrint)
                                        {
                                            emailFingerPrint += print;
                                        }
                                        if (emailDict.Keys.Contains(emailFingerPrint))
                                        {
                                            var mailList = emailDict[emailFingerPrint];
                                            if (mailList.Keys.Contains(pieceString))
                                            {
                                                mailList[pieceString]++;
                                            }
                                            else
                                            {
                                                emailDict[emailFingerPrint].Add(pieceString, 1);
                                            }
                                        }
                                        else
                                        {
                                            var newError = new Dictionary<string, int>() { { pieceString, 1 } };
                                            emailDict.Add(emailFingerPrint, newError);
                                        }
                                        //// Save email to local disk
                                        //oMail.SaveAs(@"C:\Users\Alex\Desktop\3ShapeEmailStatsGenerator\Emails\email" + k + ".txt", true);
                                    }

                                }
                                var formattedFile = new List<string>();
                                var msgFile = new List<string>();
                                emailDict = emailDict.OrderByDescending(x => x.Value.Values.Sum()).ToDictionary(x => x.Key, x => x.Value);
                                Console.WriteLine("stop");
                                formattedFile.Add("Total Unique Error Fingerprints (first three words of email): " + emailDict.Keys.Count());
                                formattedFile.Add("~~~~~~SUMMARY~~~~~");
                                foreach (var _key in emailDict.Keys)
                                {
                                    formattedFile.Add("Number of error emails for fingerprint " + _key + ": " + emailDict[_key].Values.Sum());
                                }

                                msgFile.Add("~~~~~~DATA~~~~~");
                                foreach (var err in emailDict.Keys)
                                {
                                    var _emails = emailDict[err].OrderByDescending(x => x.Value).ToDictionary(d => d.Key, d => d.Value);
                                    var _keys = _emails.Keys;
                                    var _valueSum = _emails.Values.Sum();
                                    msgFile.Add("\t" + "Total Unique Errors by Fingerprint " + err + ": " + _valueSum);
                                    foreach (var uniqueEmail in _keys)
                                    {
                                        msgFile.Add("\t\t" + "Number of Emails for the Following Entry: " + _emails[uniqueEmail]);
                                        msgFile.Add("\t\t\t" + uniqueEmail);
                                    }
                                }
                                File.WriteAllLines(@"C: \Users\Alex\Desktop\3ShapeEmailStatsGenerator\Emails\summary.txt", formattedFile);
                                File.WriteAllLines(@"C: \Users\Alex\Desktop\3ShapeEmailStatsGenerator\Emails\sortedErrorList.dat", msgFile);
                            }
                        }
                    }
                    //break;
                }

                Imap4Folder destFolder = null;
                for (int i = 0; i < count; i++)
                {
                    Imap4Folder folder = folders[i];
                    if (String.Compare("Deleted Items", folder.Name, true) == 0)
                    {
                        //find "Deleted Items" folder
                        destFolder = folder;
                        break;
                    }
                }

                if (destFolder == null)
                    throw new Exception("Deleted Items not found!");

                //MailInfo[] infos = oClient.GetMailInfos();
                //count = infos.Length;
                //for (int i = 0; i < count; i++)
                //{
                //    MailInfo info = infos[i];
                //    //move to  "Deleted Items" folder
                //    oClient.Move(info, destFolder);
                //}

                oClient.Logout();           
                //oClient.Timeout = 30;
                //oClient.Connect(oServer);
                ////var folder = oClient.Imap4Folders[0].Name;
                ////var x = 2;
                //
                    

                //    // Mark email as deleted from POP3 server.
                //    //oClient.Delete(info);
                //}

                //// Quit and purge emails marked as deleted from POP3 server.
                //oClient.Quit();
            }
            catch (Exception ep)
            {
                Console.WriteLine(ep.Message);
            }
        }
    }
}
/* var body = oMail.TextBody;
                                        string[] stringSeparators = new string[] { "\r\n" };
                                        string[] pieces = body.Split(stringSeparators, StringSplitOptions.RemoveEmptyEntries);
                                        var pieceList = pieces.ToList();
                                        var _emailFingerPrint = pieces[0].Split(' ').Take(3);
                                        var emailFingerPrint = "";
                                        foreach (var print in _emailFingerPrint)
                                        {
                                            emailFingerPrint += print;
                                        }
                                        if (emailDict.Keys.Contains(emailFingerPrint))
                                        {
                                            var mailList = emailDict[emailFingerPrint];
                                            if (mailList.Keys.Contains(pieceList))
                                            {
                                                mailList[pieceList]++;
                                            }
                                            else
                                            {
                                                emailDict[emailFingerPrint].Add(pieceList, 1);
                                            }
                                        }
                                        else
                                        {
                                            var newError = new Dictionary<List<string>, int>() { { pieceList, 1 } };
                                            emailDict.Add(emailFingerPrint, newError);
                                        }
                                        
                                    }
                                    var formattedFile = new List<string>();
                                    emailDict = emailDict.OrderBy(x => x.Value.Values.Sum()).ToDictionary(x => x.Key, x => x.Value);
                                    Console.WriteLine("stop");
                                    formattedFile.Add("Total Unique Error Fingerprints (first three words of email): " + emailDict.Keys.Count());
                                    foreach (var err in emailDict.Keys)
                                    {
                                        var _emails = emailDict[err].OrderBy(x => x.Value).ToDictionary(d => d.Key, d => d.Value);
                                        var _keys = _emails.Keys;
                                        var _valueSum = _emails.Values.Sum();
                                        formattedFile.Add("Total Unique Errors by Fingerprint " + err + ": " + (_keys.Count() + _valueSum));
                                        foreach (var uniqueEmail in _keys)
                                        {
                                            formattedFile.Add(" Number of Emails for the Following Entry: " + _emails[uniqueEmail]);
                                            foreach (var line in uniqueEmail)
                                            {
                                                formattedFile.Add("     " + line);
                                            }
                                        }
                                    }
                                    //// Save email to local disk
                                    //oMail.SaveAs(fileName, true);*/
