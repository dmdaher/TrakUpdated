
using System;
using System.Security;
using Microsoft.SharePoint.Client;
using SP = Microsoft.SharePoint.Client;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text;
using System.IO;
using EAGetMail; // add EAGetMail namespace
using Microsoft.Exchange.WebServices.Data;
using System.Configuration;
using ReadingEmailFromExchange;
using HtmlAgilityPack;
using System.Text.RegularExpressions;

namespace WindowsFormsApp1
{
    public class Program
    {
        ExchangeService serviceInstance;
        public string ExceptionMessage { get; }
        //public string MProjectTitle { get; set; }
        //public string MBody { get; set; }
        private string mProjectTitle;
        private string mBody;
        private string mDateAndTimeSent;
        private string mActualStartDate;
        private string mActualStartTime;
        //milestone arrays, pos, and size
        private string[] mMilestoneNum;
        private int mMilestoneSize;
        private int mMilestoneCurrentArrPosition;
        private string[] mMilestoneCommand;
        private string[] mMilestoneComment;
        private string mEstEndDate;
        private string mEstStartDate;
        private string mEstEndTime;
        private string mEstStartTime;
        private string mResources;
        private string mTimeSpent;
        private string mCurrentStatus;
        private string mStatusReason;

        Program()
        {
            this.mMilestoneNum = new string[20];
            this.mMilestoneCommand = new string[20];
            this.mMilestoneComment = new string[20];
            mMilestoneSize = 0;
        }
        //subject and body of email
        public string getMProjectTitle() { return mProjectTitle; }
        public void setMProjectTitle(string value) { mProjectTitle = value; }

        public string getMBody() { return mBody; }
        public void setMBody(string value) { mBody = value; }

        //metadata like date and time sent, received, etc.
        public string getMDateAndTimeSent() { return mDateAndTimeSent; }
        public void setMDateAndTimeSent(string value) { mDateAndTimeSent = value; }

        //milestone size
        public int getMMilestoneSize() { return mMilestoneSize; }
        //milestone number, command, and comment
        public string getMMilestoneNum(int pos) { return mMilestoneNum[pos]; }
        //keep milestone number around under 4-5
        public void setMMilestoneNum(string value)
        {
            int position;
            position = Convert.ToInt32(value);
            int arrPosition = position - 1;
            if(arrPosition < mMilestoneNum.Length - 1 && arrPosition >= 0)
            {
                mMilestoneNum[arrPosition] = value;
                mMilestoneSize++;
            }
            else
            {
                Console.WriteLine("Milestone Number too large!");
            }
            mMilestoneCurrentArrPosition = arrPosition;

        }

        public string getMMilestoneCommand(int pos) { return mMilestoneCommand[pos]; }
        public void setMMilestoneCommand(string value)
        {
            Console.WriteLine("we in COMMMANDD %%%%%%%%%%%");
            if (mMilestoneCurrentArrPosition < mMilestoneCommand.Length - 1)
            {
                mMilestoneCommand[mMilestoneCurrentArrPosition] = value;
            }
            else
            {
                Console.WriteLine("Milestone Command Number too large!");
            }
            Console.WriteLine("the position for command is: " + mMilestoneCurrentArrPosition);
        }

        public string getMMilestoneComment(int pos) { return mMilestoneComment[pos]; }
        public void setMMilestoneComment(string value)
        {
            Console.WriteLine("we in COMMENTSS &&&&&&&&&&&");
     
            if (mMilestoneCurrentArrPosition < mMilestoneComment.Length - 1)
            {
                mMilestoneComment[mMilestoneCurrentArrPosition] = value;
            }
            else
            {
                Console.WriteLine("Milestone Comment POosition too large!");
            }
            Console.WriteLine("the position is: " + mMilestoneCurrentArrPosition);
        }

        //estimated end date and time
        public string getMEstEndDate() { return mEstEndDate; }
        public void setMEstEndDate(string value) { mEstEndDate = value; }

        public string getMEstEndTime() { return mEstEndTime; }
        public void setMEstEndTime(string value) { mEstEndTime = value; }

        //estimated start date and time
        public string getMEstStartDate() { return mEstStartDate; }
        public void setMEstStartDate(string value) { mEstStartDate = value; }

        public string getMEstStartTime() { return mEstStartTime; }
        public void setMEstStartTime(string value) { mEstStartTime = value; }

        //actual start date and time
        public string getMActualStartTime() { return mActualStartTime; }
        public void setMActualStartTime(string value) { mActualStartTime = value; }

        public string getMActualStartDate() { return mActualStartDate; }
        public void setMActualStartDate(string value) { mActualStartDate = value; }

        //resources
        public string getMResources() { return mResources; }
        public void setMResources(string value) { mResources = value; }

        //time spent
        public string getMTimeSpent() { return mTimeSpent; }
        public void setMTimeSpent(string value) { mTimeSpent = value; }

        //currentstatus and status reason
        public string getMCurrentStatus() { return mCurrentStatus; }
        public void setMCurrentStatus(string value) { mCurrentStatus = value; }

        public string getMStatusReason() { return mStatusReason; }
        public void setMStatusReason(string value) { mStatusReason = value; }



        //<summary>
        // The main entry point for the application.
        //</summary>

        [STAThread]
        static void Main()
        {
            try
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);

                //connect to exchange
                //autodiscoverurl
                ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
                service.Credentials = new WebCredentials("devin@denaliai.com", "Welcome2018");
                service.UseDefaultCredentials = false;
                service.TraceEnabled = true;
                service.TraceFlags = TraceFlags.All;
                service.AutodiscoverUrl("devin@denaliai.com", RedirectionUrlValidationCallback);

                //write an email

                //Console.WriteLine("the contents are: " + service.Url);
                //EmailMessage email = new EmailMessage(service);
                //email.ToRecipients.Add("devin@denaliai.com");
                //email.Subject = "HelloWorld";
                //email.Body = new MessageBody("Content: This is the first email I've sent by using the EWS Managed API.");
                //email.Send();

                // Bind the Inbox folder to the service object.
                Microsoft.Exchange.WebServices.Data.Folder inbox = Microsoft.Exchange.WebServices.Data.Folder.Bind(service, WellKnownFolderName.Inbox);
                // The search filter to get unread email.
                Console.WriteLine("BEFORE");
                SearchFilter sf = new SearchFilter.SearchFilterCollection(LogicalOperator.And, new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false));

                //new SearchFilter.ContainsSubstring(EmailMessageSchema.Sender, "@denaliai.com", ContainmentMode.Substring, ComparisonMode.IgnoreCase)
                Console.WriteLine("the email message schema sender is: " + EmailMessageSchema.Sender);
                //SearchFilter.ContainsSubstring subjectFilter = new SearchFilter.ContainsSubstring(EmailMessageSchema.Sender,"@denaliai.com", ContainmentMode.Substring, ComparisonMode.IgnoreCase);

                Console.WriteLine("AFTER");

                ItemView view = new ItemView(1);
                // Fire the query for the unread items.
                // This method call results in a FindItem call to EWS.
                view.PropertySet = new PropertySet(BasePropertySet.IdOnly, EmailMessageSchema.Sender);
                
                FindItemsResults<Item> findResults = service.FindItems(WellKnownFolderName.Inbox, sf, view);
                //service.LoadPropertiesForItems(findResults, view.PropertySet);
                string body = "";
                string subject = "";
                Program prog = new Program();
                foreach (Item item in findResults.Items)
                {
                    EmailMessage message = EmailMessage.Bind(service, item.Id, PropertySet.FirstClassProperties);

                    message.Load();
                    
                    body = message.Body.Text; //MAYBE TRY TO CONVERT TO JUST THE BODY TEXT NOT THE HTML

                    subject = message.Subject;

                    string dateAndTimeSent = message.DateTimeSent.ToString();
                    prog.setMDateAndTimeSent(dateAndTimeSent);

                    Console.WriteLine("the date and time sent is: " + dateAndTimeSent);
                    Console.WriteLine("the email read is: " + subject);

                    //Do other stuff
                }
                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                doc.LoadHtml(body);
                String fullBodyText = "";
                foreach (HtmlNode node in doc.DocumentNode.SelectNodes("//text()"))
                {
                    if(node.InnerText != "")
                    {
                        Console.WriteLine(node.InnerText);
                        fullBodyText += "\n" + node.InnerText;
                    }
                }
               
                //fullBodyText = adjustString(fullBodyText);
                prog.setMBody(fullBodyText);
                prog.setMProjectTitle(subject);
                Console.WriteLine("body is: " + prog.getMBody());
                prog.parseEmail(prog);
               
                
            }
            catch(Exception e)
            {
                Console.WriteLine(e);
            }
            
        }

        //PARSE THE STRING
        //LOOK FOR KEYWORDS LIKE Milestone, Estimated Start Time, Resources,
        //get the value after keyword and put that in variable to add to SP list
        public void parseEmail(Program prog)
        {
            //Program prog = new Program();
            string aLine = null;
            StringReader strReader = new StringReader(this.getMBody());
            int actualCount = 0;
            int milestoneCount = 0;
            Console.WriteLine("the private variable is: " + this.mProjectTitle);
            aLine = strReader.ReadLine();
            while (aLine != null)
            {
                
                //aLine.ToLower();
                //aLine = aLine.Replace(" ", String.Empty);
                aLine.Trim();
                if (aLine != null && aLine != "")
                {
                    //aParagraph = aParagraph + aLine + " ";                
                    //Console.WriteLine("the line is: " + aLine);
                    string parsedValue = valueParser(aLine);
                    if (aLine.Contains("Milestone"))
                    {
                        
                        //setting milestone converts the string milestone to a milestone number
                        prog.setMMilestoneNum(parsedValue);
                        int pos = Convert.ToInt32(parsedValue);
                        Console.WriteLine("milestone is SET! " + prog.getMMilestoneNum(pos));
                        milestoneCount++;

                    }
                    //gets milestone command
                    //if command is remove, sets comment to ""
                    else if (aLine.Contains("Command"))
                    {  
                        prog.setMMilestoneCommand(parsedValue);
                        if(parsedValue == "Remove")
                        {
                            prog.setMMilestoneComment("");
                        }
                        Console.WriteLine("milestoneCommand is READY! " + prog.getMMilestoneCommand(mMilestoneCurrentArrPosition));
                    }
                    //gets milestone comment and adds it to correlated position with command and num
                    //if there is no comment like when a remove command is given, comment is set to "" in command if statement
                    else if(aLine.Contains("Comment"))
                    {                       
                        prog.setMMilestoneComment(parsedValue);
                        Console.WriteLine("milestone is PERFECT! " + prog.getMMilestoneComment(mMilestoneCurrentArrPosition));
                    }
                    //calls helper function to retrieve estimated date and times
                    else if (aLine.Contains("Estimated"))
                    {
                        helperEstimatedParser(aLine, prog);
                    }
                    //gets actual date and time
                    else if(actualCount == 0)
                    {
                        helperActualParser(prog);
                        actualCount++;
                    }
                    //gets milestone comment
                    else if(aLine.Contains("Resources"))
                    {
                        
                        prog.setMResources(parsedValue);
                        Console.WriteLine("Found the resources needed to be... " + prog.getMResources());
                    }
                    else if (aLine.Contains("Time Spent"))
                    {
                        
                        prog.setMTimeSpent(parsedValue);
                        Console.WriteLine("Found the time spent as...  " + prog.getMTimeSpent());
                    }
                    else if (aLine.Contains("Current Status"))
                    {
                        
                        prog.setMCurrentStatus(parsedValue);
                        Console.WriteLine("Found the current status as...  " + prog.getMCurrentStatus());
                    }
                    else if (aLine.Contains("Status Reason"))
                    {
                        
                        prog.setMStatusReason(parsedValue);
                        Console.WriteLine("Found the status reason as...  " + prog.getMStatusReason());
                    }
                }
                aLine = strReader.ReadLine();
            }
            Sharepoint sp = new Sharepoint(prog);
            sp.TryGetList(prog.getMProjectTitle());     
            sp.addAllData(prog.getMProjectTitle());
        }

        public void helperActualParser(Program prog)
        {

            //prog.setMActualDate(actualDateParser());
            //prog.setMActualTime(actualTimeParser());
            string dateTimeSent = prog.getMDateAndTimeSent();
            string[] dateTimeArr = new string[2];
            actualDateTimeParser(dateTimeSent, dateTimeArr);
            prog.setMActualStartDate(dateTimeArr[1]);
            prog.setMActualStartTime(dateTimeArr[0]);
            Console.WriteLine("the actual start date is: " + prog.getMActualStartDate());
            Console.WriteLine("the actual start time is: " + prog.getMActualStartTime());
        }
        public void helperEstimatedParser(string str, Program prog)
        {
            //Console.WriteLine("original string is: " + str);
            if (str.Contains("&nbsp;"))
            {
                str = str.Replace("&nbsp;", " ");
            }
            //Console.WriteLine("string after nbsp is: " + str);
            string originalStr = str;
            string parsedTime = estValueParser(str);
            string parsedDate = valueParser(originalStr);
            
            //check original string to see what title it contains
            //then input the parsed value into the correct category
            if(str.Contains("Estimated Start Time"))
            {
                prog.setMEstStartTime(parsedTime);
                Console.WriteLine("est start time has been set to:  " + prog.getMEstStartTime());
            }
            else if(str.Contains("Estimated End Time"))
            {
                prog.setMEstEndTime(parsedTime);
                Console.WriteLine("est end time has been set to:  " + prog.getMEstEndTime());
            }
            else if (str.Contains("Estimated Start Date"))
            {
                prog.setMEstStartDate(parsedDate);
                Console.WriteLine("est start date has been set to:  " + prog.getMEstStartDate());
            }
            else if (str.Contains("Estimated End Date"))
            {
                prog.setMEstEndDate(parsedDate);
                Console.WriteLine("est end date has been set to:  " + prog.getMEstEndDate());
            }
        }

        //parse the values out of the email like milestone num, command, comment, etc.
        //starts from end of the string
        public string valueParser(string str)
        {
            Console.WriteLine("finding current status reason string is: " + str);
            //str.Replace("&nbsp;", " ");
            str.Trim();
            string newStr = "";
            for (int i = str.Length - 1; i >= 0; i--)
            {
                if (str[i] == ':' || str[i] == ';' || str[i] == '#')
                {
                    break;
                }
                else
                {
                    newStr = str[i] + newStr;
                }
            }
            newStr.Trim();
            return newStr;
        }

        public string estValueParser(string str)
        {
            //str.Replace("&nbsp;", " ");
            str.Trim();
            string newStr = "";
            int colonCount = 0;
            for (int i = str.Length - 1; i >= 0; i--)
            {
                if (str[i] == ':')
                {
                    colonCount++;
                    if(colonCount >= 2) 
                    {
                        break;
                    }
                    else
                    {
                        newStr = str[i] + newStr;
                    }
                }
                else
                {
                    newStr = str[i] + newStr;
                }
            }
            newStr.Trim();
            return newStr;
        }

        //put the date and time into an array
        //array first index is the timesent
        //array second index is the datesent
        public string[] actualDateTimeParser(string str, string[] strArr)
        {
            str.Trim();
            string newTimeStr = "";
            string newDateStr = "";
            int spaceCount = 0;
            int slashCount = 0;
            for (int i = str.Length - 1; i >= 0; i--)
            {
                if (str[i] == ' ') spaceCount++;
                if (spaceCount >= 2)
                {
                    if (str[i] == '/') slashCount++;
                    newDateStr = str[i] + newDateStr;
                    Console.WriteLine("the date is: " + newDateStr);
                    if (slashCount >=2)
                    {
                        newDateStr.Trim();
                        strArr[1] = newDateStr;
                        Console.WriteLine("the date in the if is: " + newDateStr);

                    }
                }
                else
                {
                    newTimeStr = str[i] + newTimeStr;
                    newTimeStr.Trim();
                    strArr[0] = newTimeStr;
                    Console.WriteLine("the time is: " + newTimeStr);
                }
            }
            return strArr;
        }

        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            // The default for the validation callback is to reject the URL.
            bool result = false;
            Uri redirectionUri = new Uri(redirectionUrl);
            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;
        }
    }
}



//namespace Microsoft.SDK.SharePointServices.Samples
//{
//    class CreateListItem
//    {
//        static void Main()
//        {
//            string siteUrl = "https://uscedu.sharepoint.com/sites/Test123";
//            SecureString password = new SecureString();

//            foreach (char c in "Devindaher10".ToCharArray())
//                password.AppendChar(c);

//            ClientContext clientContext = new ClientContext(siteUrl);
//            SP.List oList = clientContext.Web.Lists.GetByTitle("Announcements");
//            clientContext.Credentials = new SharePointOnlineCredentials("ddaher@usc.edu", password);

//            ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
//            ListItem oListItem = oList.AddItem(itemCreateInfo); ;
//            oListItem["Title"] = "My New Item!";
//            oListItem["Body"] = "Hello World! It is a pleasure to be here" +
//                "What can I say?" +
//                "I'm simply human";

//            oListItem.Update();

//            clientContext.ExecuteQuery();
//        }
//    }
//}

