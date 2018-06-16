using Microsoft.Exchange.WebServices.Data;
using System;

namespace WindowsFormsApp1
{
    internal class EmailErrors
    {
        private Program mProg;
        private int mErrorCode;
        public EmailErrors(Program prog)
        {
            this.mProg = prog;
            this.mErrorCode = -1;
            Console.WriteLine("ALRIGHTY");
        }

        public void setErrorCode(int code) { this.mErrorCode = code; }
        public int getErrorCode() { return this.mErrorCode; }
        public void sendMilestoneTooLargeErrorEmail()
        {
            EmailMessage email = new EmailMessage(mProg.getMService());
            email.ToRecipients.Add("devin@denaliai.com");
            email.Subject = mProg.getMProjectTitle();
            email.Body = new MessageBody("OOOWWWEEEEE you really messed up that milestone number didn't you. " +
                "WAYYYY TOO LARGE. Well, larger by the biggest one by more than 2. Should be consecutive numbers, cha feel?");
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