using MailKit;
using MailKit.Net.Imap;
using MailKit.Net.Smtp;
using MimeKit;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EML2PST
{
    class mail
    {
        String sourcePath = "emails\\";
        String mAttPath = "attachments\\";
        public bool uploadMail(MimeKit.HeaderList headers, MimeKit.MimeMessage message, String StoreToFolder)
        {
            bool res = true;
            messages msg = new messages();
            var subject = headers[MimeKit.HeaderId.Subject];
            var recipient = headers[MimeKit.HeaderId.To];
            var  sender = headers[MimeKit.HeaderId.From];
            var sent = headers[MimeKit.HeaderId.Date];
            msg.msgCOMMON("Received " + subject);
            if(subject == null)
            {
                msg.msgCOMMON("Set subject to empty value");
                subject = "";
            }

            // TODO: Refactor this to store in external file
            String mFoldersListIMAP;
            String IMAPServer = ""; // Set Your email IMAP server
            int IMAPPort = 993;                 // Set IMAP port
            String MailAddress = ""; // Set email address 
            String MailName = "Test";           // Set email title
            String MailPassword = ""; // Set email password
            String SMTPServer = ""; // Set outgoing server
            int SMTPPort = 465;                 // Set outgoing port
            String InitialFolder = "INBOX";     // Set initial folder to store emails to
            String FullPath = InitialFolder;
            if (StoreToFolder.Length > 0)
            {
                FullPath += "/" + StoreToFolder;
            }
            bool folderFound = false;
            try
            {
                msg.msgCOMMON("Create connection");
                MimeKit.MimeMessage mMail = new MimeKit.MimeMessage();
                BodyBuilder mBody = new BodyBuilder();
                ImapClient mClient = new ImapClient();
                SmtpClient mSmtp = new SmtpClient();

                mClient.CheckCertificateRevocation = false;
                mClient.Connect(IMAPServer, IMAPPort);
                mClient.Authenticate(MailAddress, MailPassword); 
                mClient.Inbox.Open(FolderAccess.ReadWrite);
                msg.msgCOMMON("Create mail");

                mMail.Subject = subject;
                mBody.HtmlBody = message.HtmlBody;
                DirectoryInfo attachmentsDirectory = new DirectoryInfo(mAttPath);
                foreach (FileInfo file in attachmentsDirectory.GetFiles())
                {
                    mBody.Attachments.Add(file.FullName);
                    msg.msgCOMMON("Attach " + file.FullName);
                }


                mMail.Body = mBody.ToMessageBody();
                mSmtp.CheckCertificateRevocation = false;
                mSmtp.Connect(SMTPServer, SMTPPort, true);
                mSmtp.Authenticate(MailAddress, MailPassword);
               
                var root = mClient.GetFolder("");

                FolderNamespace fnms = new FolderNamespace('/', "");
                mFoldersListIMAP = "";

                try
                {
                    folderFound = mClient.GetFolder(FullPath).Exists;
                    
                } catch  (Exception ex)  {

                }
                
                if (folderFound)
                {
                    msg.msgGREEN("Folder exists " + FullPath);
                } else
                {
                    msg.msgGREEN("Folder not exists " + FullPath + " create");
                        var z = mClient.GetFolder(InitialFolder).Create(StoreToFolder, true);
                }


                foreach (var fldr in mClient.GetFolders(fnms))
                {
                    mFoldersListIMAP = mFoldersListIMAP + ";" + fldr.FullName;


                }
                msg.msgCOMMON("Folder list " + mFoldersListIMAP);
                

                UniqueId uid = new UniqueId();
                msg.msgCOMMON("Parse date " + sent);
                DateTimeOffset dto;
                if (!DateTimeOffset.TryParse(sent, out dto))
                    dto = DateTimeOffset.Now;
                msg.msgCOMMON("Set sent date to " + dto);
                
                mMail.Date = dto;
                msg.msgCOMMON("Parse sender ");
                InternetAddressList senderAddresses = InternetAddressList.Parse(sender);
                msg.msgCOMMON("Newreceiptient ");
                InternetAddressList recipientAddresses = new InternetAddressList();
                // TODO: Refactor this
                if (recipient != null)
                {
                    msg.msgCOMMON("Parse recipient");
                    try
                    {
                        recipientAddresses = InternetAddressList.Parse(recipient);
                    } catch (Exception ex)
                    {
                        mMail.Bcc.Add(new MailboxAddress("default@noname.com")); //set default BCC if no To was set in original mail
                    }
                } else
                {
                    msg.msgRED("Set To addr");
                    mMail.Bcc.Add(new MailboxAddress("default@noname.com"));//set default BCC if no To was set in original mail
                }
                msg.msgCOMMON("Sender: " + senderAddresses);
                msg.msgCOMMON("Recipient " + recipientAddresses);
                foreach (InternetAddress adr in senderAddresses)
                {
                    mMail.From.Add(adr);
                }
                foreach (InternetAddress adr in recipientAddresses)
                {
                    mMail.To.Add(adr);
                }

                uid = (UniqueId)mClient.GetFolder(FullPath).Append(mMail);
                mClient.GetFolder(FullPath).Open(FolderAccess.ReadWrite);
                mClient.GetFolder(FullPath).AddFlags(uid, MailKit.MessageFlags.Seen, true);
                mClient.Disconnect(true);
                mSmtp.Disconnect(true);
            }
            catch (Exception e)
            {
                msg.msgCOMMON("Exception uploading mail " +  " " + e.Message);
                res = false;
            }

            return res;
        }

        

    }
}
