using Microsoft.Exchange.WebServices.Data;
using System;

namespace WindowsFormsApp1
{
    public class EmailErrors
    {
        private Program mProg;
        private int mErrorCode;
        private bool errorCheck;
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
        public void sendMilestoneTooLargeErrorEmail()
        {
            EmailMessage email = new EmailMessage(mProg.getMService());
            email.ToRecipients.Add("devin@denaliai.com");
            email.Subject = mProg.getMProjectTitle();
            email.Body = new MessageBody("OOOWWWEEEEE you really messed up that milestone number didn't you. " +
                "WAYYYY TOO LARGE. Well, larger by the biggest one by more than 2. Should be consecutive numbers, cha feel?" +
                "Or you just created a new list without starting at 1 with your milestone. lol");
            email.Send();
        }
        public void sendMilestoneAlreadyExistsErrorEmail()
        {
            EmailMessage email = new EmailMessage(mProg.getMService());
            email.ToRecipients.Add("devin@denaliai.com");
            email.Subject = mProg.getMProjectTitle();
            email.Body = new MessageBody("Haha the milestone already exists! What were you thinking? Try again and resend a valid milestone");
            email.Send();
        }
        public void sendMilestoneErrorEmail()
        {
            EmailMessage email = new EmailMessage(mProg.getMService());
            email.ToRecipients.Add("devin@denaliai.com");
            email.Subject = mProg.getMProjectTitle();
            email.Body = new MessageBody("OOOWWWEEEEE you really messed up that milestone number didn't you");
            email.Send();
        }
    }
}