
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
    public class MailItem
    {
        public string mFrom;
        public string[] mRecipients;
        public string mSubject;
        public string mBody;
        public string mDateAndTimeSent;
        public string mMailboxAddress;
        public string mMailboxName;
        public bool mIsAuthority;

        public MailItem()
        {
            mIsAuthority = false;
            mSubject = "";
        }

        public void loadMail(FindItemsResults<Item> findResults, ExchangeService service, Program prog)
        {
            foreach (Item item in findResults.Items)
            {
                EmailMessage message = EmailMessage.Bind(service, item.Id, PropertySet.FirstClassProperties);
                message.Load();
                this.mFrom = message.From.ToString();
                this.mMailboxAddress = message.From.Address;
                this.mMailboxName = message.From.Name;

                //if(mail.mMailboxAddress == "omar.hassan@t-mobile.com" || mail.mMailboxName == "Omar Hassan") { mail.mIsAuthority = true; }
                if (this.mMailboxAddress == "ddaher@usc.edu") { this.mIsAuthority = true; }
                    //if (this.mMailboxAddress == "ddaher@usc.edu")
                    //{
                    //    EmailErrors ee = new EmailErrors(prog);
                    //    ee.sendMilestoneErrorEmail();
                    //}
                    //string[] recipients = ((Microsoft.Exchange.WebServices.Data.EmailAddressCollection)item[EmailMessageSchema.ToRecipients]).Select(recipient => recipient.Address).ToArray();
                Console.WriteLine("the address is: " + this.mFrom);
                Console.WriteLine("the mailbox type is: " + message.From.MailboxType);
                Console.WriteLine("the mailbox address is: " + message.From.Address);
                Console.WriteLine("the mailbox name is: " + message.From.Name);
                this.mBody = message.Body.Text;
                this.mSubject = message.Subject;


                

                string dateAndTimeSent = message.DateTimeSent.ToString();
                this.mDateAndTimeSent = message.DateTimeSent.ToString();
                prog.setMDateAndTimeSent(dateAndTimeSent);

                Console.WriteLine("the date and time sent is: " + this.mDateAndTimeSent);
                
                Console.WriteLine("the email read is: " + this.mSubject);

                //Do other stuff
            }
        }
        
    }

    public class Milestone
    {
        string command;
        string comment;
        int number;

        public Milestone(int num, string command, string comment)
        {
            this.command = command;
            this.comment = comment;
            this.number = num;
        }
        public void setCommand(string command) { this.command = command; }
        public string getCommand() { return this.command; }
        public void setComment(string comment) { this.comment = comment; }
        public string getComment() { return this.comment; }
        public void setNumber(int num) { this.number = num; }
        public int getNumber() { return this.number; }
    }
    public class Program
    {
        private EmailErrors mEmailError;
        private ExchangeService mService;
        public string ExceptionMessage { get; }
        //public string MProjectTitle { get; set; }
        //public string MBody { get; set; }
        private string mProjectTitle;
        private string mBody;
        private string mDateAndTimeSent;
        private string mActualStartDate;
        private string mActualStartTime;
        //milestone arrays, pos, and size
        private Milestone mCurrMilestone;
        //private string[] mMilestoneNum;
        private Dictionary<int, string> mMilestoneNumMap;
        private Dictionary<int, string> mMilestoneCommandMap;
        private Dictionary<int, string> mMilestoneCommentMap;
        private Dictionary<int, Milestone> mMilestoneObjMap;
        private int mMilestoneSize;
        private int mMilestoneCurrentNum;
        //private string[] mMilestoneCommand;
        //private string[] mMilestoneComment;
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
            Service service = new Service();
            this.mService = service.getMService();
            //this.mMilestoneNum = new string[20];
            //this.mMilestoneCommand = new string[20];
            //this.mMilestoneComment = new string[20];
            this.mCurrMilestone = new Milestone(0, "", "");
            this.mMilestoneObjMap = new Dictionary<int, Milestone>();
            this.mMilestoneNumMap = new Dictionary<int, string>();
            this.mMilestoneCommandMap = new Dictionary<int, string>();
            this.mMilestoneCommentMap = new Dictionary<int, string>();
            mMilestoneSize = 0;
            this.mEmailError = new EmailErrors(this);
            
        }

        //service instance
        public ExchangeService getMService() { return mService; }
        //subject and body of email
        public string getMProjectTitle() { return mProjectTitle; }
        public void setMProjectTitle(string value) { mProjectTitle = value; }

        public string getMBody() { return mBody; }
        public void setMBody(string value) { mBody = value; }

        //metadata like date and time sent, received, etc.
        public string getMDateAndTimeSent() { return mDateAndTimeSent; }
        public void setMDateAndTimeSent(string value) { mDateAndTimeSent = value; }

        //milestone object
        public void setMilestone()
        {
            try
            {
                Console.WriteLine("inside try ");
                int num = this.mCurrMilestone.getNumber();
                if (!mMilestoneObjMap.TryGetValue(num, out Milestone result))
                {
                    Console.WriteLine("the key when setting milestone is: " + num);
                    this.mMilestoneObjMap.Add(num, this.mCurrMilestone);
                    Console.WriteLine("added key: " + num);
                }
                else
                {
                    Console.WriteLine("milestone already exists...Are you trying to replace it?!");
                    //this.mEmailError.sendMilestoneErrorEmail();  //won't get into this else bc creates a new program everytime
                    //need to figure out how to maintain same program OR read data that already exists first and check with new data inputted
                }
            }
            catch (ArgumentNullException ane)
            {
                Console.WriteLine("setting milestone caught exception" + ane);
            }
            catch (ArgumentException ae)
            {
                Console.WriteLine("setting milestone caught exception #2" + ae);
            }
        }
        public Milestone getMilestone(int pos)
        {
            mMilestoneObjMap.TryGetValue(pos, out Milestone value);
            if (value == null)
            {
                value = null;
            }
            //mMilestoneNumMap.TryGetValue(pos, out value);
            Console.WriteLine("NUM: what is the value num if the key is found? " + value.getNumber());
            Console.WriteLine("NUM: what is the value comment if the key is found " + value.getComment());
            return value;
        }
        public Dictionary<int, Milestone> getMilestoneObjMap()
        {
            return this.mMilestoneObjMap;
        }
        //milestone size
        public int getMMilestoneSize() { return mMilestoneSize; }

        //grab milestone number dictionary
        public Dictionary<int, string> getMMilestoneNumMap()
        {
            return this.mMilestoneNumMap;
        }
        public Dictionary<int, string> getMMilestoneCommandMap()
        {
            return this.mMilestoneCommandMap;
        }
        public Dictionary<int, string> getMMilestoneCommentMap()
        {
            return this.mMilestoneCommentMap;
        }
        //milestone number, command, and comment
        public string getMMilestoneNum(int pos)
        {
            //fix get function, use conditional statement if possible. looks cleaner
            //string value = "";
            mMilestoneNumMap.TryGetValue(pos, out string value);
            if (value == null)
            {
                value = "";
            }
            //mMilestoneNumMap.TryGetValue(pos, out value);
            Console.WriteLine("NUM: what is the value if the key is NOT found? " + value);
            Console.WriteLine("NUM: what is the value if the key is found " + value); 
            return value;
        }
        //keep milestone number around under 4-5
        public void setMMilestoneNum(string value)
        {
            
            int key;
            key = Convert.ToInt32(value);
            Console.WriteLine("the key when setting milestone is: " + key);

            //string result = "";
            try
            {
                Console.WriteLine("inside try ");
                if (!mMilestoneNumMap.TryGetValue(key, out string result))
                {
                    Console.WriteLine("the key when setting milestone is: " + key);
                    mMilestoneNumMap.Add(key, value);
                    Console.WriteLine("added key: " + key);
                }
                else
                {
                    Console.WriteLine("key already exists...Are you trying to replace it?!");
                    //this.mEmailError.sendMilestoneErrorEmail();  //won't get into this else bc creates a new program everytime
                    //need to figure out how to maintain same program OR read data that already exists first and check with new data inputted
                }
            }
            catch(ArgumentNullException ane)
            {
                Console.WriteLine("setting milestone caught exception" + ane);
            }catch(ArgumentException ae)
            {
                Console.WriteLine("setting milestone caught exception #2" + ae);
            }
            
            mMilestoneSize++;
            mMilestoneCurrentNum = key;
            //Console.WriteLine("the position when setting milestone is: " + value);
            //int arrPosition = key - 1;
            //if(arrPosition < mMilestoneNum.Length - 1 && arrPosition >= 0)
            //{
            //    mMilestoneNum[arrPosition] = value;
            //    mMilestoneSize++;
            //    //Console.WriteLine("in if setting milestone");
            //}
            //else
            //{
            //    Console.WriteLine("Milestone Number too large!");
            //}


        }

        public string getMMilestoneCommand(int pos)
        {          
            mMilestoneCommandMap.TryGetValue(pos, out string value);
            if (value == null)
            {
                value = "";
            }
            Console.WriteLine("COMMAND: what is the value if the key is NOT found? " + value + " and position is: " + pos);
            //Console.WriteLine("COMMAND: what is the value if the key is found " + value);
            return value;
            //return mMilestoneCommand[pos];
        }
        public void setMMilestoneCommand(string value)
        {
            //int key;
            //key = Convert.ToInt32(value);
            //mMilestoneCommandMap.Add(key, value); //what if key already exists & what if invalid key like a string
            Console.WriteLine("we in COMMMANDD %%%%%%%%%%%");

            Console.WriteLine("the command key when setting milestone is: " + value);

            //string result = "";
            try
            {
                Console.WriteLine("inside command try ");
                if (!mMilestoneCommandMap.TryGetValue(mMilestoneCurrentNum, out string result))
                {
                    Console.WriteLine("the key when setting milestone is: " + value);
                    mMilestoneCommandMap.Add(mMilestoneCurrentNum, value);
                    Console.WriteLine("added command key: " + value);
                }
                else
                {
                    Console.WriteLine("command key already exists...Are you trying to replace it?!");
                }
            }
            catch (ArgumentNullException ane)
            {
                Console.WriteLine("setting milestone caught exception" + ane);
            }
            catch (ArgumentException ae)
            {
                Console.WriteLine("setting milestone caught exception #2" + ae);
            }
            //if (mMilestoneCurrentArrPosition < mMilestoneCommand.Length - 1)
            //{
            //    mMilestoneCommand[mMilestoneCurrentArrPosition] = value;
            //}
            //else
            //{
            //    Console.WriteLine("Milestone Command Number too large!");
            //}
            //Console.WriteLine("the position for command is: " + mMilestoneCurrentArrPosition);
        }

        public string getMMilestoneComment(int pos)
        {
            string value = "";
            mMilestoneCommentMap.TryGetValue(pos, out value);
            if(value == null)
            {
                value = "";
            }
            Console.WriteLine("COMMENT: what is the value if the comment key is NOT found? " + value);
            Console.WriteLine("Comment: what is the value if the comment key is found " + value);
            return value;
            //return mMilestoneComment[pos];
        }
        public void setMMilestoneComment(string value)
        {
            Console.WriteLine("current milestone pos is: " + mMilestoneCurrentNum);
            Console.WriteLine("we in COMMENTSS &&&&&&&&&&&");
            //string result = "";
            try
            {
                Console.WriteLine("inside comment try ");
                if (!mMilestoneCommentMap.TryGetValue(mMilestoneCurrentNum, out string result))
                {
                    Console.WriteLine("the key when setting milestone is: " + mMilestoneCurrentNum);
                    mMilestoneCommentMap.Add(mMilestoneCurrentNum, value);
                    Console.WriteLine("added comment key: " + value);
                }
                else
                {
                    Console.WriteLine("comment key already exists...Are you trying to replace it?!");
                }
            }
            catch (ArgumentNullException ane)
            {
                Console.WriteLine("setting milestone caught exception" + ane);
            }
            catch (ArgumentException ae)
            {
                Console.WriteLine("setting milestone caught exception #2" + ae);
            }
            //if (mMilestoneCurrentArrPosition < mMilestoneComment.Length - 1)
            //{
            //    mMilestoneComment[mMilestoneCurrentArrPosition] = value;
            //}
            //else
            //{
            //    Console.WriteLine("Milestone Comment POosition too large!");
            //}
            //Console.WriteLine("the position is: " + mMilestoneCurrentArrPosition);
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




        //private ExchangeService mExchService;

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
                //Service service = new Service();
                Program prog = new Program();
                ExchangeService service = prog.getMService();          

                // Bind the Inbox folder to the service object.
                Microsoft.Exchange.WebServices.Data.Folder inbox = Microsoft.Exchange.WebServices.Data.Folder.Bind(service, WellKnownFolderName.Inbox);
                // The search filter to get unread email.
                Console.WriteLine("BEFORE");

                //TIP to create search filter
                //first create search collection list
                //add any search filter like a filter that looks for a certain email or checks if email is unread or read
                //then if you want more searches, create a new search filter that ANDS the previous list collection
                //then if you want to let's say check for a domain name in the email, create new list
                //with that new list, add new search filter & add previous filter
                //then create a new search filter and AND or OR it with previous collection
                //idea is this starts chaining filters together
                //Another tip: Make sure you properly use AND and OR...e.g. can't filter two email domains 
                //  and say you want the substring to contain denali AND tmobile
                //  so you choose OR operator to say i want to filter for either domain

                //SearchFilter sf = new SearchFilter.SearchFilterCollection(LogicalOperator.And, new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false));
                List<SearchFilter> searchANDFilter = new List<SearchFilter>();
                //searchANDFilter.Add(sf);
                ExtendedPropertyDefinition PidTagSenderSmtpAddress = new ExtendedPropertyDefinition(0x5D01, MapiPropertyType.String);
                //searchANDFilter.Add(new SearchFilter.ContainsSubstring(PidTagSenderSmtpAddress, "@denaliai.com"));
                searchANDFilter.Add(new SearchFilter.ContainsSubstring(PidTagSenderSmtpAddress, "@denaliai.com"));
                searchANDFilter.Add(new SearchFilter.ContainsSubstring(PidTagSenderSmtpAddress, "@usc.edu"));
                //Console.WriteLine("the address is: " + PidTagSenderSmtpAddress.PropertySet.Value.ToString());
                SearchFilter domainSF = new SearchFilter.SearchFilterCollection(LogicalOperator.Or, searchANDFilter);
                List<SearchFilter> searchFinalFilter = new List<SearchFilter>();
                searchFinalFilter.Add(new SearchFilter.SearchFilterCollection(LogicalOperator.And, new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false)));
                searchFinalFilter.Add(domainSF);
                SearchFilter finalSF = new SearchFilter.SearchFilterCollection(LogicalOperator.And, searchFinalFilter);
                //new SearchFilter.ContainsSubstring(EmailMessageSchema.Sender, "@denaliai.com", ContainmentMode.Substring, ComparisonMode.IgnoreCase)
                //SearchFilter.ContainsSubstring subjectFilter = new SearchFilter.ContainsSubstring(EmailMessageSchema.Sender,"@denaliai.com", ContainmentMode.Substring, ComparisonMode.IgnoreCase);

                Console.WriteLine("AFTER");

                ItemView view = new ItemView(1);
                // Fire the query for the unread items.
                // This method call results in a FindItem call to EWS.
                view.PropertySet = new PropertySet(BasePropertySet.IdOnly, EmailMessageSchema.Sender);
                
                //FindItemsResults<Item> findResults = service.FindItems(WellKnownFolderName.Inbox, sf, view);
                FindItemsResults<Item> findResults = service.FindItems(WellKnownFolderName.Inbox, finalSF, view);
                //service.LoadPropertiesForItems(findResults, view.PropertySet);
                //Program prog = new Program();

                try
                {
                    MailItem mail = new MailItem();
                    mail.loadMail(findResults, service, prog);
                    prog.setMProjectTitle(mail.mSubject);
                    Console.WriteLine("&&&&&&&&&&&&&&&&&&&&&&&&the project title is: " + prog.getMProjectTitle());
                    if (mail.mIsAuthority == true)
                    { //idea is to send project report if authority is true
                        //EmailErrors ee = new EmailErrors(prog);
                        //ee.sendMilestoneErrorEmail();
                        Sharepoint sp = new Sharepoint(prog);
                        string fullReport = sp.createProjectReport(prog.getMProjectTitle());
                    }
                    else
                    {
                        HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                        //doc.LoadHtml(body);
                        doc.LoadHtml(mail.mBody);
                        String fullBodyText = "";
                        foreach (HtmlNode node in doc.DocumentNode.SelectNodes("//text()"))
                        {
                            if (node.InnerText != "")
                            {
                                Console.WriteLine(node.InnerText);
                                fullBodyText += "\n" + node.InnerText;
                            }
                        }
                        //fullBodyText = adjustString(fullBodyText);
                        prog.setMBody(fullBodyText);
                        //prog.setMProjectTitle(subject);
                        prog.setMProjectTitle(mail.mSubject);
                        Console.WriteLine("body is: " + prog.getMBody());
                        prog.parseEmail(prog, service);
                    }
                }catch(Exception e)
                {
                    Console.WriteLine(e);
                } 
            }
            catch(Exception e)
            {
                Console.WriteLine(e);
            }
            
        }

        //PARSE THE STRING
        //LOOK FOR KEYWORDS LIKE Milestone, Estimated Start Time, Resources,
        //get the value after keyword and put that in variable to add to SP list
        public void parseEmail(Program prog, ExchangeService service)
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
                    parsedValue = parsedValue.Trim();
                    if (aLine.Contains("Milestone"))
                    {
                        this.mCurrMilestone = new Milestone(0, "","");
                        //setting milestone converts the string milestone to a milestone number
                        prog.setMMilestoneNum(parsedValue);
                        int pos = Convert.ToInt32(parsedValue);
                        this.mCurrMilestone.setNumber(pos);
                        Console.WriteLine("milestone is SET! " + prog.getMMilestoneNum(pos));
                        milestoneCount++;
                        //mMilestoneCurrentNum = pos;
                    }
                    //gets milestone command
                    //if command is remove, sets comment to ""
                    else if (aLine.Contains("Command"))
                    {
                        prog.setMMilestoneCommand(parsedValue);
                        this.mCurrMilestone.setCommand(parsedValue);
                        if (parsedValue == "Remove")
                        {
                            prog.setMMilestoneComment("");
                            this.mCurrMilestone.setComment("");
                            this.setMilestone();
                        }
                        Console.WriteLine("milestoneCommand is READY! " + prog.getMMilestoneCommand(mMilestoneCurrentNum));
                    }
                    //gets milestone comment and adds it to correlated position with command and num
                    //if there is no comment like when a remove command is given, comment is set to "" in command if statement
                    else if(aLine.Contains("Comment"))
                    {                       
                        prog.setMMilestoneComment(parsedValue);
                        Console.WriteLine("milestone is PERFECT! " + prog.getMMilestoneComment(mMilestoneCurrentNum));
                        this.mCurrMilestone.setComment(parsedValue);
                        this.setMilestone();
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
            try
            {
                bool isNewList = false;
                EmailErrors emailError = sp.getMEmailErrors();
                isNewList = sp.TryGetList(prog.getMProjectTitle());
                sp.addAllData(prog.getMProjectTitle());
                //if (!isNewList)
                //{
                //    sp.readData(prog.getMProjectTitle());
                //}
                
                //Create Response Email
                //EmailMessage email = new EmailMessage(service);
                //email.ToRecipients.Add("devin@denaliai.com");
                //email.Subject = "Updated Project: " + mProjectTitle;
                //email.Body = new MessageBody("Your Sharepoint project has successfully been updated! Looking forward to your next update :)");
                //email.Send();


                
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            
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
            str = str.Replace("&nbsp;", " ").Trim();
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
    }
}


//NOTE: when inputting password, you must add to char array
//SecureString password = new SecureString();

//            foreach (char c in "Devindaher10".ToCharArray())
//                password.AppendChar(c);

//            ClientContext clientContext = new ClientContext(siteUrl);
//            SP.List oList = clientContext.Web.Lists.GetByTitle("Announcements");
//            clientContext.Credentials = new SharePointOnlineCredentials("ddaher@usc.edu", password);

