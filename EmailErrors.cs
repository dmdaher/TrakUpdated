using Microsoft.Exchange.WebServices.Data;
using System;

namespace WindowsFormsApp1
{
    public class EmailErrors
    {
        private Program mProg;
        private int mErrorCode;
        private bool errorCheck;
        private string mailboxAddress;
        public EmailErrors(Program prog)
        {
            this.errorCheck = false;
            this.mProg = prog;
            this.mErrorCode = -1;
            Console.WriteLine("ALRIGHTY");
        }

        public void setErrorCheck(bool value) { this.errorCheck = value; }
        public bool getErrorCheck() { return this.errorCheck; }
        public void setErrorCode(int code) { this.mErrorCode = code; }
        public int getErrorCode() { return this.mErrorCode; }
        public void sendMilestoneTooLargeErrorEmail(string mailboxAddress)
        {
            EmailMessage email = new EmailMessage(mProg.getMService());
            //email.ToRecipients.Add("devin@denaliai.com");
            email.ToRecipients.Add(mailboxAddress);
            email.Subject = mProg.getMProjectTitle();
            email.Body = new MessageBody("OOOWWWEEEEE you really messed up that milestone number didn't you. " +
                "WAYYYY TOO LARGE. Well, larger by the biggest one by more than 2. Should be consecutive numbers, cha feel?" +
                "Or you just created a new list without starting at 1 with your milestone. lol");
            email.Send();
        }
        public void sendMilestoneAlreadyExistsErrorEmail(string mailboxAddress)
        {
            EmailMessage email = new EmailMessage(mProg.getMService());
            //email.ToRecipients.Add("devin@denaliai.com");
            email.ToRecipients.Add(mailboxAddress);
            email.Subject = mProg.getMProjectTitle();
            email.Body = new MessageBody("Haha the milestone already exists! What were you thinking? Try again and resend a valid milestone");
            email.Send();
        }
        public void sendMilestoneErrorEmail(string mailboxAddress)
        {
            EmailMessage email = new EmailMessage(mProg.getMService());
            //email.ToRecipients.Add("devin@denaliai.com");
            email.ToRecipients.Add(mailboxAddress);
            email.Subject = mProg.getMProjectTitle();
            email.Body = new MessageBody("OOOWWWEEEEE you really messed up that milestone number didn't you");
            email.Send();
        }

        public void sendProjectDoesNotExistEmail(string mailboxAddress)
        {
            EmailMessage email = new EmailMessage(mProg.getMService());
            email.ToRecipients.Add(mailboxAddress);
            email.Subject = mProg.getMProjectTitle();
            email.Body = new MessageBody("Project does not exist in Share Point. Please try again with a new project title");
            email.Send();
        }
    }
}