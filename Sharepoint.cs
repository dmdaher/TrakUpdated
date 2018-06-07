using System;
using System.Security;
using Microsoft.SharePoint.Client;
using SP = Microsoft.SharePoint.Client;
using System.Linq;
using System.Collections.Generic;

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

        public void addList()
        {
            //To refactor code, in order to garbage collect, maybe use this using statement to be safer
            //using (ClientContext context = new ClientContext("http://yourserver/"))
            //{
            //    context.Credentials = new NetworkCredential("user", "password", "domain");
            //    List list = context.Web.Lists.GetByTitle("Some List");
            //    context.ExecuteQuery();

            //    // Now update the list.
            //}


            // Starting with ClientContext, the constructor requires a URL to the 
            // server running SharePoint.  

            // The SharePoint web at the URL.
            Web web = mClientContext.Web;
            if(mClientContext.Web.Lists.GetByTitle(mProg.getMProjectTitle()) == null)
            {
                ListCreationInformation creationInfo = new ListCreationInformation();
                creationInfo.Title = mProg.getMProjectTitle();
                creationInfo.TemplateType = (int)ListTemplateType.GenericList;
                List list = web.Lists.Add(creationInfo);
                list.Description = "Project Updates";
                list.Update();
                mClientContext.ExecuteQuery();
                addColumns(); //add static columns for each list created
            }
        }

        public void addColumns()
        {
            Web web = mClientContext.Web;
            SP.List projList = mClientContext.Web.Lists.GetByTitle(mProg.getMProjectTitle());
            if (projList != null)
            {
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
                //SP.Field oField = collFields.GetByTitle("MyNewField");
                Console.WriteLine("Executing end of adding columns");
                mClientContext.ExecuteQuery();
            }
        }

        public void addEstimatedTimesColumn()
        {
            Web web = mClientContext.Web;
            SP.List oList = mClientContext.Web.Lists.GetByTitle(mProg.getMProjectTitle());
            ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
            ListItem oListItem = oList.AddItem(itemCreateInfo);
            //oListItem["Title"] = mProg.getMEstStartDate();
            oListItem["Title"] = "ALRIGHTY";
            //oListItem["Body"] = mProg.getMEstStartTime();
            Console.WriteLine("MUST UPDATE");
            oListItem.Update();

            mClientContext.ExecuteQuery();
        }
    }
    //SPList newList = web.Lists["My List"];

    //// create Text type new column called "My Column"
    //newList.Fields.Add("My Column", SPFieldType.Text, true);
    //        SP.List oList = mClientContext.Web.Lists.GetByTitle("Announcements");
    //ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
    //ListItem oListItem = oList.AddItem(itemCreateInfo); ;
    //        oListItem["Title"] = "My New Item!";
    //        oListItem["Body"] = "Hello World! It is a pleasure to be here" +
    //            "What can I say?" +
    //            "I'm simply human";

    //        oListItem.Update();

    //        clientContext.ExecuteQuery();
}

//SYNTAX:
//to grab field //SP.Field oField = collFields.GetByTitle("MyNewField");
//