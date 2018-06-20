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
        private int mInitialMilestoneRowCount;
        private List<int> milestonesWithAddCommands;
        public Sharepoint(Program prog)
        {
            this.mIsNewList = true;
            this.mInitialMilestoneRowCount = 0;
            milestonesWithAddCommands = new List<int>();
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

        public EmailErrors getMEmailErrors()
        {
            return this.mEmailError;
        }

        public void setMIsNewList(bool value)
        {
            this.mIsNewList = value;
        }

        public bool TryGetList(string listTitle)
        {
            Web web = mClientContext.Web;
   
            SP.List projList = web.Lists.GetByTitle(listTitle);
            try
            {
                mClientContext.ExecuteQuery();
                this.mIsNewList = false;
                return false;
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
                        this.mIsNewList = true;
                        return true;
                    }
                    return true;
                }
                catch (PropertyOrFieldNotInitializedException pfnie)
                {
                    Console.WriteLine(pfnie);
                    return true;
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
            collFields.AddFieldAsXml("<Field DisplayName='Time Spent' Type='Choice' />", true, AddFieldOptions.DefaultValue);
            collFields.AddFieldAsXml("<Field DisplayName='Resources' Type='Choice' />", true, AddFieldOptions.DefaultValue);
            collFields.AddFieldAsXml("<Field DisplayName='Current Status' Type='Choice' />", true, AddFieldOptions.DefaultValue);
            collFields.AddFieldAsXml("<Field DisplayName='Current Status Reason' Type='Choice' />", true, AddFieldOptions.DefaultValue);
            //test field
            //collFields.AddFieldAsXml("<Field DisplayName='test' Type='Choice' />", true, AddFieldOptions.DefaultValue);
            //SP.Field oField = collFields.GetByTitle("MyNewField");
            Console.WriteLine("Executing end of adding columns");
            projList.Update();
            mClientContext.ExecuteQuery();
        }


        //public void readData(string listTitle)
        //{
        //    try
        //    {
        //        Web web = mClientContext.Web;
        //        SP.CamlQuery myQuery = new SP.CamlQuery();
        //        SP.List projList = mClientContext.Web.Lists.GetByTitle(listTitle);
        //        //myQuery.ViewXml = " < Where > < Eq > < FieldRef Name = 'Estimated_x0020_Start_x0020_Time' /> <Value Type = 'Text'> 9:30 </Value> </ Eq > </ Where ></ View >";

        //        //myQuery.ViewXml = " < Where > < Eq > < FieldRef Name = 'Milestone_x0020_Number' /> <Value Type = 'Text'> 2 </Value> </ Eq > </ Where ></ View >";
        //        SP.ListItemCollection collectItems = projList.GetItems(myQuery);
        //        //SP.ListItemCollection collectItems = projList.GetItems(CamlQuery.CreateAllItemsQuery());
        //        mClientContext.Load(collectItems);
        //        mClientContext.ExecuteQuery();
        //        //Dictionary<int, string> milestoneDictionary = mProg.getMMilestoneNumMap();
        //        //Dictionary<int, string> milestoneCommandDict = mProg.getMMilestoneCommandMap();
        //        //Dictionary<int, string> milestoneCommentDict = mProg.getMMilestoneCommentMap();
        //        int count = collectItems.Count;
        //        Dictionary<int, Milestone> milestoneDict = mProg.getMilestoneObjMap();
        //        bool errorCheck = this.mEmailError.getErrorCheck();
        //        int sizeOfMap = mProg.getMilestoneObjMap().Count;
        //        int maxInputKey = 0;
        //        if(milestoneDict.Count > 0)
        //        {
        //            maxInputKey = milestoneDict.Keys.Max();
        //            Console.WriteLine("largest key in map before entering for loop is: " + maxInputKey);
        //        }

        //        int maxNumInTable = 0;

        //        //error code = -1 means default, no error has occurred so far
        //        //error code = 0 means success, continue
        //        //error code = 1 is milestone already exists
        //        //error code = 2 is milestone is too large
        //        //so far that is the system
        //        foreach (SP.ListItem oItem in collectItems)
        //        {
        //            string milestoneStrInTable = (string)oItem["Milestone_x0020_Number"];
        //            int milestoneNumInTable = Convert.ToInt32(milestoneStrInTable);
        //            Console.WriteLine("key in table is: " + milestoneNumInTable);
        //            if (milestoneNumInTable > maxNumInTable)
        //            {
        //                maxNumInTable = milestoneNumInTable;
        //            }
        //            if (milestoneDict.ContainsKey(milestoneNumInTable))
        //            {
        //                //break when error code is not default and not success -- means error occurred
        //                if (this.mEmailError.getErrorCode() != 0 && this.mEmailError.getErrorCode() != -1)
        //                {
        //                    break;
        //                }
        //                else
        //                {
        //                    milestoneDict.TryGetValue(milestoneNumInTable, out Milestone currMilestone);  //result holds the milestone that is correlated with the key
        //                    string command = currMilestone.getCommand().Trim();
        //                    Console.WriteLine("value of command is: " + command);
        //                    if (command == "Add") { this.mEmailError.setErrorCheck(true); }
        //                    milestoneCommandHandler(projList, oItem, currMilestone);
        //                }
        //                //Error duplicate key, send email back
        //            }
        //            //need to check if table doesn't have key b/c user may input milestone to update that does not exist
        //        }
        //        //ADD THIS ERROR BACK IN IN IN ININI!!!!!!!!

        //        //if(maxInputKey != maxNumInTable + 1 && this.mIsNewList == false && this.mEmailError.getErrorCode() == -1)
        //        //{
        //        //    this.mEmailError.setErrorCheck(true);
        //        //    this.mEmailError.setErrorCode(2);
        //        //    Console.WriteLine("ERROR 2 --- TOO LARGE");
        //        //    Console.WriteLine("what is the max input key? " + maxInputKey + " and the max in table is: " + maxNumInTable);
        //        //    mEmailError.sendMilestoneTooLargeErrorEmail();
        //        //    //Error, too large of a milestone inputted
        //        //}
        //        if (!this.mEmailError.getErrorCheck())
        //        {
        //            mClientContext.ExecuteQuery();
        //            //addAllData(listTitle);
        //        }
        //    }catch(CollectionNotInitializedException cnie)
        //    {
        //        if (cnie.Message.Contains("The collection has not been initialized"))
        //        {
        //            Console.WriteLine("WEL WLE EWLL");
        //        }

        //    }catch(InvalidOperationException ioe)
        //    {
        //        Console.WriteLine(ioe);
        //    }

        //}

        //this is the read data that does not update anything. it simply reads and outputs error if anything wrong is happening
        //like an already existing milestone
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
                int count = collectItems.Count;
                Dictionary<int, Milestone> milestoneDict = mProg.getMilestoneObjMap();
                bool errorCheck = this.mEmailError.getErrorCheck();
                int sizeOfMap = mProg.getMilestoneObjMap().Count;
                int minInputKey = 0;
                if (milestoneDict.Count > 0)
                {
                    minInputKey = milestoneDict.Keys.Min();
                    Console.WriteLine("largest key in map before entering for loop is: " + minInputKey);
                }

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
                    milestoneDict.TryGetValue(milestoneNumInTable, out Milestone currMilestone);  //result holds the milestone that is correlated with the key
                    if(currMilestone != null)
                    {
                        string command = currMilestone.getCommand().Trim();
                        if (milestoneNumInTable > maxNumInTable)
                        {
                            maxNumInTable = milestoneNumInTable;
                        }
                        if (milestoneDict.ContainsKey(milestoneNumInTable))
                        {
                            //break when error code is not default and not success -- means error occurred
                            if (this.mEmailError.getErrorCode() != 0 && this.mEmailError.getErrorCode() != -1)
                            {
                                break;
                            }
                            else
                            {
                                if (command == "Add")
                                {
                                    this.mEmailError.setErrorCheck(true);
                                }
                                milestoneCommandHandler(projList, oItem, currMilestone);
                                //Console.WriteLine("value of command is: " + command);
                                //if (command == "Add")
                                //{
                                //    this.mEmailError.setErrorCheck(true);
                                //    mEmailError.sendMilestoneAlreadyExistsErrorEmail();
                                //    this.mEmailError.setErrorCode(1);
                                //}
                                //else if (command == "Update")
                                //{
                                //    oItem["Milestone_x0020_Comment"] = currMilestone.getComment();
                                //    this.mEmailError.setErrorCode(0);
                                //}
                                //else if (command == "Remove")
                                //{
                                //    oItem["Milestone_x0020_Comment"] = "";
                                //    this.mEmailError.setErrorCode(0);
                                //}
                                //oItem.Update();
                                
                            }
                            //Error duplicate key, send email back
                        }
                        //else
                        //{
                        //    if (command == "Add")
                        //    {
                        //        this.milestonesWithAddCommands.Add(currMilestone.getNumber());
                        //    }
                        //}
                    }
                    //need to check if table doesn't have key b/c user may input milestone to update that does not exist
                }
                //ADD THIS ERROR BACK IN IN IN ININI!!!!!!!!

                if (minInputKey > maxNumInTable + 1 && this.mIsNewList == false && this.mEmailError.getErrorCode() == -1)
                {
                    this.mEmailError.setErrorCheck(true);
                    this.mEmailError.setErrorCode(2);
                    Console.WriteLine("ERROR 2 --- TOO LARGE");
                    Console.WriteLine("what is the max input key? " + minInputKey + " and the max in table is: " + maxNumInTable);
                    mEmailError.sendMilestoneTooLargeErrorEmail();
                    //Error, too large of a milestone inputted
                }
                if (!this.mEmailError.getErrorCheck())
                {
                    mClientContext.ExecuteQuery();
                    //addAllData(listTitle);
                }
            }
            catch (CollectionNotInitializedException cnie)
            {
                if (cnie.Message.Contains("The collection has not been initialized"))
                {
                    Console.WriteLine("WEL WLE EWLL");
                }

            }
            catch (InvalidOperationException ioe)
            {
                Console.WriteLine(ioe);
            }

        }

        public void milestoneCommandHandler(List projList, ListItem oItem, Milestone currMilestone)
        {
            string command = currMilestone.getCommand().Trim();
            string comment = currMilestone.getComment().Trim();
            int number = currMilestone.getNumber();
            if (command == "Add") //adding milestone that already exists. ERROR
            {
                if (this.mEmailError.getErrorCheck())
                {
                    this.mEmailError.setErrorCheck(true);
                    mEmailError.sendMilestoneAlreadyExistsErrorEmail();
                    this.mEmailError.setErrorCode(1);
                }
                else
                {
                    //explain why i have this counter?..
                    if (this.mInitialMilestoneRowCount != 0)
                    {
                        Console.WriteLine("in Add if statement");
                        ListItemCreationInformation itemCreateMilestone = new ListItemCreationInformation();
                        ListItem milestoneListItem = projList.AddItem(itemCreateMilestone);
                        //int milestoneNum = Convert.ToInt32(currMilestone.getNumber());
                        Console.WriteLine("milestone integer number converted now is: " + number);
                        milestoneListItem["Milestone_x0020_Number"] = number;
                        milestoneListItem["Milestone_x0020_Comment"] = comment;
                        milestoneListItem.Update();
                    }
                    else
                    {
                        Console.WriteLine("milestone integer number converted now is: " + number);
                        oItem["Milestone_x0020_Number"] = number;
                        oItem["Milestone_x0020_Comment"] = comment;
                        //oItem.Update();
                        this.mInitialMilestoneRowCount++;
                    }
                }
            }
            //remove milestone - just removes the comment
            //sets errors code to 0 meaning success
            else if (command == "Remove") //removing milestone comment
            {
                Console.WriteLine("MUST CHANGE COMMENT & REMOVE! -- the key is: " + number);
                oItem["Milestone_x0020_Comment"] = "";
                this.mEmailError.setErrorCode(0);
            }
            else if (command == "Update") //updating milestone
            {
                string newComment = comment;
                oItem["Milestone_x0020_Comment"] = newComment;
                this.mEmailError.setErrorCode(0);
            }
            oItem.Update();
        }

        public void milestoneAddHandler(List projList, ListItem oItem, Milestone currMilestone)
        {
            string command = currMilestone.getCommand().Trim();
            string comment = currMilestone.getComment().Trim();
            int number = currMilestone.getNumber();
            if (command == "Add")
            {
                if (this.mEmailError.getErrorCheck())
                {
                    this.mEmailError.setErrorCheck(true);
                    mEmailError.sendMilestoneAlreadyExistsErrorEmail();
                    this.mEmailError.setErrorCode(1);
                }
                else
                {
                    //explain why i have this counter?..
                    if (this.mInitialMilestoneRowCount != 0)
                    {
                        Console.WriteLine("in Add if statement");
                        ListItemCreationInformation itemCreateMilestone = new ListItemCreationInformation();
                        ListItem milestoneListItem = projList.AddItem(itemCreateMilestone);
                        //int milestoneNum = Convert.ToInt32(currMilestone.getNumber());
                        Console.WriteLine("milestone integer number converted now is: " + number);
                        milestoneListItem["Milestone_x0020_Number"] = number;
                        milestoneListItem["Milestone_x0020_Comment"] = comment;
                        milestoneListItem.Update();
                    }
                    else
                    {
                        Console.WriteLine("milestone integer number converted now is: " + number);
                        oItem["Milestone_x0020_Number"] = number;
                        oItem["Milestone_x0020_Comment"] = comment;
                        //oItem.Update();
                        this.mInitialMilestoneRowCount++;
                    }
                }
            }
            oItem.Update();
        }

        public bool milestonesAreOrdered(Dictionary<int, Milestone> milestoneDict, bool isNewList)
        {
            int milestoneNum = 0;
            int prevNum = 0;
            int initialCount = 0;
            if(milestoneDict.Count != 0)
            {
                if (milestoneDict.Keys.Min() != 1 && isNewList == true)
                {
                    return false;
                }
                foreach (var milestone in milestoneDict.OrderBy(i => i.Key))
                {
                    Milestone currMilestone = milestone.Value;
                    if(currMilestone.getCommand() == "Add")
                    {
                        milestoneNum = currMilestone.getNumber();
                        if (milestoneNum != prevNum + 1 && initialCount != 0)
                        {
                            return false;
                        }
                        Console.WriteLine("the milestone order is milestonenum with: " + milestoneNum + " and prevnum with: " + prevNum);
                        prevNum = currMilestone.getNumber();
                        initialCount++;
                    }
                }
            }
            return true;
        }

        public void inputMilestoneAddCommands(Dictionary<int, Milestone> milestoneDict)
        {
            foreach (var milestone in milestoneDict.OrderBy(i => i.Key))
            {
                Milestone currMilestone = milestone.Value;
                if(currMilestone.getCommand().Trim() == "Add")
                {
                    this.milestonesWithAddCommands.Add(currMilestone.getNumber());
                }
            }
        }
        public void addAllData(string listTitle)
        {
            Web web = mClientContext.Web;
            SP.List projList = mClientContext.Web.Lists.GetByTitle(listTitle);
            Dictionary<int, Milestone> milestoneDict = mProg.getMilestoneObjMap();
            bool isOrdered = false;
            Console.WriteLine("the error code is: " + this.mEmailError.getErrorCode());
            if (this.mIsNewList == false)
            {
                readData(listTitle);
                isOrdered = milestonesAreOrdered(milestoneDict, false);
                inputMilestoneAddCommands(milestoneDict);
            }
            else
            {
                isOrdered = milestonesAreOrdered(milestoneDict, true);
                inputMilestoneAddCommands(milestoneDict);
            }
            if ((this.mEmailError.getErrorCode() == 0 || this.mEmailError.getErrorCode() == -1) && isOrdered == true)
            {
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
                List<int> addCommand = new List<int>();
                List<int> updateCommand = new List<int>();
                List<int> removeCommand = new List<int>();
                //if(this.mIsNewList == true)
                //{
                //    foreach (var pair in mProg.getMilestoneObjMap().OrderBy(i => i.Key))
                //    {
                //        Milestone milestone = pair.Value;
                //        Console.WriteLine("the milestone command is: " + milestone.getCommand().Trim());
                //        if (milestone.getCommand().Trim() == "Add")
                //        {
                //            milestoneAddHandler(projList, oListItem, milestone);
                //        }
                //    }
                //}
                if(this.milestonesWithAddCommands.Count > 0)
                {
                    for (int i = 0; i < this.milestonesWithAddCommands.Count; i++)
                    {
                        milestoneDict.TryGetValue(this.milestonesWithAddCommands[i], out Milestone milestone);
                        if (milestone != null)
                        {
                            milestoneAddHandler(projList, oListItem, milestone);
                            //if (milestone.getCommand().Trim() == "Update")
                            //{
                            //    updateCommand.Add(milestone.getNumber());
                            //}
                            //if (milestone.getCommand().Trim() == "Remove")
                            //{
                            //    removeCommand.Add(milestone.getNumber());
                            //}
                        }
                    }
                }

                //foreach (var pair in mProg.getMilestoneObjMap().OrderBy(i => i.Key))
                //{
                //    Milestone milestone = pair.Value;
                //    Console.WriteLine("the milestone command is: " + milestone.getCommand().Trim());
                //    if (milestone.getCommand().Trim() == "Add")
                //    {
                //        addCommand.Add(milestone.getNumber());
                //        milestoneCommandHandler(projList, oListItem, milestone);
                //    }
                //    else if (milestone.getCommand().Trim() == "Update")
                //    {
                //        updateCommand.Add(milestone.getNumber());
                //    }
                //    else if (milestone.getCommand().Trim() == "Remove")
                //    {
                //        removeCommand.Add(milestone.getNumber());
                //    }
                //}
                //updateMilestones(projList, updateCommand, removeCommand, addCommand);    
                //Console.WriteLine("the milestone number is: " + mProg.getMMilestoneNum(pair.Key));
                //}
                //}else if(milestone.getCommand().Trim() == "Update" || milestone.getCommand().Trim() == "Remove")
                //{s
                //    milestoneCommandHandler(projList, oListItem, milestone);
                //}
                //oListItem.Update();
                //mClientContext.ExecuteQuery();



                //oListItem.Update();
                //SP.CamlQuery myQuery = new SP.CamlQuery();
                //myQuery.ViewXml = " < Where > < Eq > < FieldRef Name = 'Milestone_x0020_Number' /> <Value Type = 'Text'> 2 </Value> </ Eq > </ Where ></ View >";
                //SP.ListItemCollection collectItems = projList.GetItems(myQuery);
                //mClientContext.Load(collectItems);

                Console.WriteLine("MUST UPDATE");
                this.mInitialMilestoneRowCount = 0;
                oListItem.Update();

                mClientContext.ExecuteQuery();
            }
            else
            {
                this.mEmailError.setErrorCheck(true);
                if (isOrdered == false)
                {
                    mEmailError.sendMilestoneTooLargeErrorEmail();
                    this.mEmailError.setErrorCode(3);
                }
            }
        }


        public string createProjectReport(string listTitle)
        {
            try
            {
                Web web = mClientContext.Web;
                SP.CamlQuery myQuery = new SP.CamlQuery();
                SP.List projList = mClientContext.Web.Lists.GetByTitle(listTitle);
                //myQuery.ViewXml = " < Where > < Eq > < FieldRef Name = 'Estimated_x0020_Start_x0020_Time' /> <Value Type = 'Text'> 9:30 </Value> </ Eq > </ Where ></ View >";

                //myQuery.ViewXml = " < Where > < Eq > < FieldRef Name = 'Milestone_x0020_Number' /> <Value Type = 'Text'> 2 </Value> </ Eq > </ Where ></ View >";
                //SP.ListItemCollection collectItems = projList.GetItems(myQuery);
                //myQuery.ViewXml = "< OrderBy > < FieldRef Name = 'Title' Ascending = 'False' /> </ OrderBy >";
                SP.ListItemCollection collectItems = projList.GetItems(CamlQuery.CreateAllItemsQuery());
                
                mClientContext.Load(collectItems);
                mClientContext.ExecuteQuery();
                int count = collectItems.Count;
                string fullReport = "";
                foreach (SP.ListItem oItem in collectItems)
                {
                    string estimatedStartTime = (string)oItem["Estimated_x0020_Start_x0020_Time"];
                    string estimatedEndTime = (string)oItem["Estimated_x0020_End_x0020_Time"];
                    string estimatedStartDate = (string)oItem["Estimated_x0020_Start_x0020_Date"];
                    string estimatedEndDate = (string)oItem["Estimated_x0020_End_x0020_Date"];
                    string actualStartTime = (string)oItem["Actual_x0020_Start_x0020_Time"];
                    string actualEndTime = (string)oItem["Actual_x0020_End_x0020_Time"];
                    string actualStartDate = (string)oItem["Actual_x0020_Start_x0020_Date"];
                    string actualEndDate = (string)oItem["Actual_x0020_End_x0020_Date"];
                    string milestoneNumber = (string)oItem["Milestone_x0020_Number"];
                    string milestoneComment = (string)oItem["Milestone_x0020_Comment"];
                    string timeSpent = (string)oItem["Time_x0020_Spent"];
                    string resources = (string)oItem["Resources"];
                    string currentStatus = (string)oItem["Current_x0020_Status"];
                    string currentStatusReason = (string)oItem["Current_x0020_Status_x0020_Reaso"];
                    fullReport += "Here is all the data for this project: " + estimatedStartTime + "  ===  " + estimatedEndTime + "  ===  " + estimatedStartDate + "  ====  "
                        + estimatedEndDate + "  ===  " + milestoneNumber + " ===  " + milestoneComment+ " \n /n";
                    Console.WriteLine(fullReport);
                }
                return fullReport;
            }
            catch (CollectionNotInitializedException cnie)
            {
                if (cnie.Message.Contains("The collection has not been initialized"))
                {
                    Console.WriteLine(cnie);
                }
                return "";
            }
            catch (InvalidOperationException ioe)
            {
                Console.WriteLine(ioe);
                return "";
            }
        }
    }
}

//SYNTAX:
//to grab field //SP.Field oField = collFields.GetByTitle("MyNewField");
//Note : If you have space in your display name it's converted to '_x0020_' , '_' is converted to '%5f'
//TIP: go to sharepoint list settings then click on column to edit then look at end of url ---- that is the internal name