using System;
using System.Security;
using Microsoft.SharePoint.Client;
using SP = Microsoft.SharePoint.Client;
using System.Linq;
using System.Collections.Generic;
using Microsoft.SharePoint;
using System.Xml;

namespace WindowsFormsApp1
{
    public class Sharepoint
    {
        private Program mProg;
        private ClientContext mClientContext;
        private SharePointOnlineCredentials mSPCreds;
        private bool mIsNewList;
        //private int mErrorCode;
        private EmailErrors mEmailError;
        public Sharepoint(Program prog)
        {
            mIsNewList = false;
            //this.mErrorCode = 0;
            this.mProg = prog;
            this.mEmailError = new EmailErrors(mProg);
            string siteUrl = "https://3mdtech.sharepoint.com/DDTest";
            SecureString password = new SecureString();

            foreach (char c in "Welcome2018".ToCharArray())
                password.AppendChar(c);

            // Starting with ClientContext, the constructor requires a URL to the 
            // server running SharePoint. 
            ClientContext clientContext = new ClientContext(siteUrl);
            mClientContext = clientContext;

            mSPCreds = new SharePointOnlineCredentials("devin@denaliai.com", password);
            mClientContext.Credentials = mSPCreds;
            //clientContext.Credentials = new SharePointOnlineCredentials("devin@denaliai.com", password);
        }

        public void TryGetList(string listTitle)
        {
            Web web = mClientContext.Web;
   
            SP.List projList = web.Lists.GetByTitle(listTitle);
            try
            {
                mClientContext.ExecuteQuery();
                mIsNewList = false;
            }
            catch (ServerException se)
            {
                try
                {
                    Console.WriteLine("the hash code is: " + se.GetHashCode());
                    if (se.Message.Contains("does not exist at site with URL"))
                    {
                        //incorrect way of checking if proj list is null or not
                        Console.WriteLine("So it is not null??");
                        ListCreationInformation creationInfo = new ListCreationInformation();
                        creationInfo.Title = listTitle;
                        creationInfo.TemplateType = (int)ListTemplateType.GenericList;
                        projList = web.Lists.Add(creationInfo);
                        projList.Description = "Project Updates";
                        projList.Update();
                        mClientContext.ExecuteQuery();
                        mIsNewList = true;
                        addColumns(listTitle);
                    }                  
                }
                catch (PropertyOrFieldNotInitializedException pfnie)
                {
                    Console.WriteLine(pfnie);
                }
            }
        }

        public void addColumns(string listTitle)
        {
            Web web = mClientContext.Web;
            SP.List projList = mClientContext.Web.Lists.GetByTitle(listTitle);
            // Adding the Custom field to the List. Here the OOB SPFieldText has been selected as the “FieldType”
            SP.FieldCollection collFields = projList.Fields;
            collFields.AddFieldAsXml("<Field DisplayName='Estimated Start Time' Type='Choice' />", true, AddFieldOptions.DefaultValue);
            collFields.AddFieldAsXml("<Field DisplayName='Estimated Start Date' Type='Choice' />", true, AddFieldOptions.DefaultValue);
            collFields.AddFieldAsXml("<Field DisplayName='Estimated End Time' Type='Choice' />", true, AddFieldOptions.DefaultValue);
            collFields.AddFieldAsXml("<Field DisplayName='Estimated End Date' Type='Choice' />", true, AddFieldOptions.DefaultValue);
            collFields.AddFieldAsXml("<Field DisplayName='Actual Start Time' Type='Choice' />", true, AddFieldOptions.DefaultValue);
            collFields.AddFieldAsXml("<Field DisplayName='Actual End Time' Type='Choice' />", true, AddFieldOptions.DefaultValue);
            collFields.AddFieldAsXml("<Field DisplayName='Actual Start Date' Type='Choice' />", true, AddFieldOptions.DefaultValue);
            collFields.AddFieldAsXml("<Field DisplayName='Actual End Date' Type='Choice' />", true, AddFieldOptions.DefaultValue);
            collFields.AddFieldAsXml("<Field DisplayName='Milestone Number' Type='Choice' />", true, AddFieldOptions.DefaultValue);
            collFields.AddFieldAsXml("<Field DisplayName='Milestone Comment' Type='Choice' />", true, AddFieldOptions.DefaultValue);
            collFields.AddFieldAsXml("<Field DisplayName='Milestone Comment' Type='Choice' />", true, AddFieldOptions.DefaultValue);
            collFields.AddFieldAsXml("<Field DisplayName='Time Spent' Type='Choice' />", true, AddFieldOptions.DefaultValue);
            collFields.AddFieldAsXml("<Field DisplayName='Resources' Type='Choice' />", true, AddFieldOptions.DefaultValue);
            collFields.AddFieldAsXml("<Field DisplayName='Current Status' Type='Choice' />", true, AddFieldOptions.DefaultValue);
            collFields.AddFieldAsXml("<Field DisplayName='Current Status Reason' Type='Choice' />", true, AddFieldOptions.DefaultValue);
            //test field
            collFields.AddFieldAsXml("<Field DisplayName='test' Type='Choice' />", true, AddFieldOptions.DefaultValue);
            //SP.Field oField = collFields.GetByTitle("MyNewField");
            Console.WriteLine("Executing end of adding columns");
            projList.Update();
            mClientContext.ExecuteQuery();
        }

        public void readData(string listTitle)
        {
            try
            {
                Web web = mClientContext.Web;
                SP.CamlQuery myQuery = new SP.CamlQuery();
                SP.List projList = mClientContext.Web.Lists.GetByTitle(listTitle);
                //myQuery.ViewXml = " < Where > < Eq > < FieldRef Name = 'Estimated_x0020_Start_x0020_Time' /> <Value Type = 'Text'> 9:30 </Value> </ Eq > </ Where ></ View >";
                
                //myQuery.ViewXml = " < Where > < Eq > < FieldRef Name = 'Milestone_x0020_Number' /> <Value Type = 'Text'> 2 </Value> </ Eq > </ Where ></ View >";
                SP.ListItemCollection collectItems = projList.GetItems(myQuery);
                //SP.ListItemCollection collectItems = projList.GetItems(CamlQuery.CreateAllItemsQuery());
                mClientContext.Load(collectItems);
                mClientContext.ExecuteQuery();
                //Dictionary<int, string> milestoneDictionary = mProg.getMMilestoneNumMap();
                //Dictionary<int, string> milestoneCommandDict = mProg.getMMilestoneCommandMap();
                //Dictionary<int, string> milestoneCommentDict = mProg.getMMilestoneCommentMap();
                int count = collectItems.Count;
                Dictionary<int, Milestone> milestoneDict = mProg.getMilestoneObjMap();
                bool errorCheck = false;
                int sizeOfMap = mProg.getMilestoneObjMap().Count;
                int maxInputKey = milestoneDict.Keys.Max();
                Console.WriteLine("largest key in map before entering for loop is: " + maxInputKey);
                int maxNumInTable = 0;
                
                //error code = -1 means default, no error has occurred so far
                //error code = 0 means success, continue
                //error code = 1 is milestone already exists
                //error code = 2 is milestone is too large
                //so far that is the system
                foreach (SP.ListItem oItem in collectItems)
                {
                    string milestoneStrInTable = (string)oItem["Milestone_x0020_Number"];
                    int milestoneNumInTable = Convert.ToInt32(milestoneStrInTable);
                    Console.WriteLine("key in table is: " + milestoneNumInTable);
                    if (milestoneNumInTable > maxNumInTable)
                    {
                        maxNumInTable = milestoneNumInTable;
                    }
                    if (milestoneDict.ContainsKey(milestoneNumInTable))
                    {                 
                        milestoneDict.TryGetValue(milestoneNumInTable, out Milestone currMilestone);  //result holds the milestone that is correlated with the key
                        string command = currMilestone.getCommand().Trim();
                        Console.WriteLine("value of command is: " + command);
                        //Error duplicate key, send email back
                        if (command == "Add")
                        {

                            Console.WriteLine("Key is duplicated! and the max in table is: " + maxNumInTable);
                            errorCheck = true;
                            mEmailError.sendMilestoneAlreadyExistsErrorEmail();
                            this.mEmailError.setErrorCode(1);
                            break;
                        }
                        //remove milestone - just removes the comment
                        //sets errors code to 0 meaning success
                        else if(command == "Remove")
                        {
                            Console.WriteLine("MUST CHANGE COMMENT & REMOVE! -- the key is: " + milestoneNumInTable);
                            oItem["Milestone_x0020_Comment"] = "";
                            this.mEmailError.setErrorCode(0);
                        }
                        else if(command == "Update")
                        {
                            string newComment = currMilestone.getComment();
                            oItem["Milestone_x0020_Comment"] = newComment;
                            this.mEmailError.setErrorCode(0);
                        }
                        oItem.Update();
                    }
                }
                if(maxInputKey != maxNumInTable + 1 && this.mIsNewList == false && this.mEmailError.getErrorCode() == -1)
                {
                    errorCheck = true;
                    this.mEmailError.setErrorCode(2);
                    Console.WriteLine("ERROR 2 --- TOO LARGE");
                    Console.WriteLine("what is the max input key? " + maxInputKey + " and the max in table is: " + maxNumInTable);
                    mEmailError.sendMilestoneTooLargeErrorEmail();
                    //Error, too large of a milestone inputted
                }
                if (!errorCheck)
                {
                    mClientContext.ExecuteQuery();
                    addAllData(listTitle);
                }
            }catch(CollectionNotInitializedException cnie)
            {
                if (cnie.Message.Contains("The collection has not been initialized"))
                {
                    Console.WriteLine("WEL WLE EWLL");
                }
                    
            }catch(InvalidOperationException ioe)
            {
                Console.WriteLine(ioe);
            }
            
        }
        public void addAllData(string listTitle)
        {
            Web web = mClientContext.Web;
            SP.List projList = mClientContext.Web.Lists.GetByTitle(listTitle);
            
            ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
            ListItem oListItem = projList.AddItem(itemCreateInfo);

            Console.WriteLine("milestone size is: " + mProg.getMMilestoneSize());
            oListItem["Estimated_x0020_Start_x0020_Time"] = mProg.getMEstStartTime();
            oListItem["Estimated_x0020_Start_x0020_Date"] = mProg.getMEstStartDate();
            oListItem["Estimated_x0020_End_x0020_Time"] = mProg.getMEstEndTime();
            oListItem["Estimated_x0020_End_x0020_Date"] = mProg.getMEstEndDate();
            oListItem["Actual_x0020_Start_x0020_Time"] = mProg.getMActualStartTime();
            oListItem["Actual_x0020_Start_x0020_Date"] = mProg.getMActualStartDate();
            oListItem["Time_x0020_Spent"] = mProg.getMTimeSpent();
            oListItem["Resources"] = mProg.getMResources();
            oListItem["Current_x0020_Status"] = mProg.getMCurrentStatus();
            oListItem["Current_x0020_Status_x0020_Reaso"] = mProg.getMStatusReason();


            SP.CamlQuery myQuery = new SP.CamlQuery();
            myQuery.ViewXml = " < Where > < Eq > < FieldRef Name = 'Milestone_x0020_Number' /> <Value Type = 'Text'> 2 </Value> </ Eq > </ Where ></ View >";
            SP.ListItemCollection collectItems = projList.GetItems(myQuery);
            mClientContext.Load(collectItems);

            mClientContext.ExecuteQuery();
            int initialMilestoneRowCount = 0;
            
            foreach (var pair in mProg.getMMilestoneNumMap())
            {
                Console.WriteLine("the milestone command is: " + mProg.getMMilestoneCommand(pair.Key).Trim());
                if (mProg.getMMilestoneCommand(pair.Key).Trim() == "Add")
                {
                    //explain why i have this counter?..
                    if (initialMilestoneRowCount != 0)
                    {
                        Console.WriteLine("in Add if statement");
                        ListItemCreationInformation itemCreateMilestone = new ListItemCreationInformation();
                        ListItem milestoneListItem = projList.AddItem(itemCreateMilestone);
                        int milestoneNum = Convert.ToInt32(mProg.getMMilestoneNum(pair.Key));
                        Console.WriteLine("milestone integer number converted now is: " + milestoneNum);
                        milestoneListItem["Milestone_x0020_Number"] = milestoneNum;
                        milestoneListItem["Milestone_x0020_Comment"] = mProg.getMMilestoneComment(pair.Key);
                        milestoneListItem.Update();
                    }
                    else
                    {
                        int milestoneNum = Convert.ToInt32(mProg.getMMilestoneNum(pair.Key));
                        Console.WriteLine("milestone integer number converted now is: " + milestoneNum);
                        oListItem["Milestone_x0020_Number"] = milestoneNum;
                        oListItem["Milestone_x0020_Comment"] = mProg.getMMilestoneComment(pair.Key);
                        oListItem.Update();
                        initialMilestoneRowCount++;
                    }
                }
                Console.WriteLine("the milestone number is: " + mProg.getMMilestoneNum(pair.Key));
                mClientContext.ExecuteQuery();
            }
            Console.WriteLine("MUST UPDATE");

            oListItem.Update();

            mClientContext.ExecuteQuery();
        }
    }
}

//SYNTAX:
//to grab field //SP.Field oField = collFields.GetByTitle("MyNewField");
//Note : If you have space in your display name it's converted to '_x0020_' , '_' is converted to '%5f'
//TIP: go to sharepoint list settings then click on column to edit then look at end of url ---- that is the internal name