using System;
using System.Security;
using Microsoft.SharePoint.Client;
using SP = Microsoft.SharePoint.Client;
using System.Linq;
using System.Collections.Generic;
using Microsoft.SharePoint;

namespace WindowsFormsApp1
{
    public class Sharepoint
    {
        private Program mProg;
        private ClientContext mClientContext;
        private SharePointOnlineCredentials mSPCreds;
        public Sharepoint(Program prog)
        {
            this.mProg = prog;
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
            projList.Fields.AddFieldAsXml("<Field DisplayName='Estimated Start Time' Type='Choice' />", true, AddFieldOptions.DefaultValue);
            projList.Fields.AddFieldAsXml("<Field DisplayName='Estimated Start Date' Type='Choice' />", true, AddFieldOptions.DefaultValue);
            projList.Fields.AddFieldAsXml("<Field DisplayName='Estimated End Time' Type='Choice' />", true, AddFieldOptions.DefaultValue);
            projList.Fields.AddFieldAsXml("<Field DisplayName='Estimated End Date' Type='Choice' />", true, AddFieldOptions.DefaultValue);
            projList.Fields.AddFieldAsXml("<Field DisplayName='Actual Start Time' Type='Choice' />", true, AddFieldOptions.DefaultValue);
            projList.Fields.AddFieldAsXml("<Field DisplayName='Actual End Time' Type='Choice' />", true, AddFieldOptions.DefaultValue);
            projList.Fields.AddFieldAsXml("<Field DisplayName='Actual Start Date' Type='Choice' />", true, AddFieldOptions.DefaultValue);
            projList.Fields.AddFieldAsXml("<Field DisplayName='Actual End Date' Type='Choice' />", true, AddFieldOptions.DefaultValue);
            projList.Fields.AddFieldAsXml("<Field DisplayName='Milestone Number' Type='Choice' />", true, AddFieldOptions.DefaultValue);
            projList.Fields.AddFieldAsXml("<Field DisplayName='Milestone Comment' Type='Choice' />", true, AddFieldOptions.DefaultValue);
            projList.Fields.AddFieldAsXml("<Field DisplayName='Time Spent' Type='Choice' />", true, AddFieldOptions.DefaultValue);
            projList.Fields.AddFieldAsXml("<Field DisplayName='Resources' Type='Choice' />", true, AddFieldOptions.DefaultValue);
            projList.Fields.AddFieldAsXml("<Field DisplayName='Current Status' Type='Choice' />", true, AddFieldOptions.DefaultValue);
            projList.Fields.AddFieldAsXml("<Field DisplayName='Current Status Reason' Type='Choice' />", true, AddFieldOptions.DefaultValue);
            //test field
            projList.Fields.AddFieldAsXml("<Field DisplayName='test' Type='Choice' />", true, AddFieldOptions.DefaultValue);
            //SP.Field oField = collFields.GetByTitle("MyNewField");
            Console.WriteLine("Executing end of adding columns");
            projList.Update();
            mClientContext.ExecuteQuery();
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
            //CamlQuery camlQuery = new CamlQuery();
            //camlQuery.ViewXml = "<View><Query><Where><Eq><FieldRef Name='Title'>" +
            //    "<Value Type='Number'>0</Value></Geq></Where></Query><RowLimit>20</RowLimit></View>";
            //ListItemCollection collListItem = projList.GetItems(camlQuery);

            //mClientContext.Load(collListItem);

            //mClientContext.ExecuteQuery();

            //foreach (ListItem queriedListItem in collListItem)
            //{
            //    Console.WriteLine("ID: {0} \nTitle: {1} \nNumber: {2}", oListItem.Id, queriedListItem["Milestone_x0020_Number"]);
            //}

            //ON THE RIGHT TRACK FOR XML QUERY FOR UPDATE AND REMOVE
            //SP.CamlQuery myQuery = new SP.CamlQuery();
            //myQuery.ViewXml = "< View >< Where >< Eq >< FieldRef Name = 'Title' /></ Eq > </ Where ></ View >";
            //  SP.ListItemCollection myItems = projList.GetItems(myQuery);
            //mClientContext.Load(myItems);
            //mClientContext.ExecuteQuery();
            //string result = string.Empty;
            //foreach(ListItem eachItem in myItems)
            //{
            //    result += eachItem["Title"].ToString() + Environment.NewLine;
            //    Console.WriteLine("result is: " + result);

            //}

            //make sure they cannot add a greater milestone number than the size
            int initialMilestoneRowCount = 0;
            //ITERATE THROUGH MAP NOT FOR LOOP
            //foreach(var pair in mProg.getMMilestoneNum())
            for (int i = 0; i<mProg.getMMilestoneSize(); i++) // foreach var in map
            {

                Console.WriteLine("the milestone command is: " + mProg.getMMilestoneCommand(i).Trim());
                
                //if (mProg.getMMilestoneCommand(i).Trim() == "Update")
                //{
                //    int milestoneNum = Convert.ToInt32(mProg.getMMilestoneNum(i));
                //    oListItem["Milestone_x0020_Number"] = projList.GetItemById(i);
                //    oListItem["Milestone_x0020_Comment"] = projList.GetItemById(i);
                //    oListItem["Milestone_x0020_Number"] = milestoneNum;
                //    oListItem["Milestone_x0020_Comment"] = mProg.getMMilestoneComment(i);
                //}
                if (mProg.getMMilestoneCommand(i).Trim() == "Add")
                {
                    if (initialMilestoneRowCount != 0)
                    {
                        Console.WriteLine("in Add if statement");
                        ListItemCreationInformation itemCreateMilestone = new ListItemCreationInformation();
                        ListItem milestoneListItem = projList.AddItem(itemCreateMilestone);
                        int milestoneNum = Convert.ToInt32(mProg.getMMilestoneNum(i));
                        Console.WriteLine("milestone integer number converted now is: " + milestoneNum);
                        milestoneListItem["Milestone_x0020_Number"] = milestoneNum;
                        milestoneListItem["Milestone_x0020_Comment"] = mProg.getMMilestoneComment(i);
                        milestoneListItem.Update();
                    }
                    else
                    {
                        int milestoneNum = Convert.ToInt32(mProg.getMMilestoneNum(i));
                        Console.WriteLine("milestone integer number converted now is: " + milestoneNum);
                        oListItem["Milestone_x0020_Number"] = milestoneNum;
                        oListItem["Milestone_x0020_Comment"] = mProg.getMMilestoneComment(i);
                        oListItem.Update();
                        initialMilestoneRowCount++;
                    }
                }
                
                //else if (mProg.getMMilestoneCommand(i).Trim() == "Remove")
                //{
                //    ListItem milestoneListItem = projList.GetItemById(mProg.getMMilestoneNum(i));
                //    milestoneListItem.DeleteObject();
                //}

                Console.WriteLine("the milestone number is: " + mProg.getMMilestoneNum(i));
                mClientContext.ExecuteQuery();
            }
            //oListItem["Milestone Number"] = mProg.getMMilestoneNum();
            
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