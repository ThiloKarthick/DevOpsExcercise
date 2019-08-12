using Microsoft.SharePoint.Client;
using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace ConsoleApp2
{
    class Program
    {
        public static void Main(string[] args)
        {
            Program p = new Program();
            //exportSQLdataToExcel();
          // p.deletelistitmsbycondition();
         //   onSAFTaskComplete.TaskOutcome = CompleteTaskStatus;
          //  p.insertistems();
            printhelloworld();
        }
        private void printhelloworld()
        {
                    Console.WriteLine("Hello World");
        }
        private void exportSQLdataToExcel()
        {
            SqlConnection cvvCon = null;
            cvvCon = new SqlConnection(@"Data Source= dca-db-432\HDSQLPRD;Initial Catalog=HDDM;Persist Security Info=True;User ID=PPVCCVVUser;Password=ppv@123#$");
            //SAP
            //Changed the Schema - Prabhakar
            SqlCommand cmdSelect = new SqlCommand("select [MATERIAL_NUMBER],[MATERIAL_DESC],[PBG],[MMPP],[plantcd],[SSA] from [CMD].[Cvv_Items] where MATERIAL_NUMBER = ''  or MATERIAL_DESC = '' or PBG = '' or plantcd = '' or MMPP = '' or SSA = '' except select top 2097150 MATERIAL_NUMBER,MATERIAL_DESC,PBG,plantcd,MMPP,SSA  FROM[HDDM].[CMD].[cvv_items] where MATERIAL_NUMBER = '' or MATERIAL_DESC = '' or PBG = '' or plantcd = '' or MMPP = '' or SSA = ''", cvvCon);
            //  SqlCommand cmdSelect = new SqlCommand("select top 1048575 [MATERIAL_NUMBER],[MATERIAL_DESC],[PBG],[MMPP],[plantcd],[SSA] from [CMD].[Cvv_Items] where MATERIAL_NUMBER = ''  or MATERIAL_DESC = '' or PBG = '' or plantcd = '' or MMPP = '' or SSA = ''", cvvCon);
            cmdSelect.CommandTimeout = 0;
            SqlDataAdapter daCVVSAP = new SqlDataAdapter(cmdSelect);
            DataTable dtSAP = new DataTable();
            daCVVSAP.Fill(dtSAP);
            //Create an instance of ExcelEngine
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                //Initialize Application
                IApplication application = excelEngine.Excel;

                //Set the default application version as Excel 2016
                application.DefaultVersion = ExcelVersion.Excel2016;

                //Create a new workbook
                IWorkbook workbook = application.Workbooks.Create(1);

                //Access first worksheet from the workbook instance
                IWorksheet worksheet = workbook.Worksheets[0];

                //Export DataTable to worksheet

                //Export DataTable to worksheet with column name and start range.
                worksheet.ImportDataTable(dtSAP, true, 1, 1);
                worksheet.UsedRange.AutofitColumns();

                //Save the workbook to disk in xlsx format
                workbook.SaveAs("CVVempty3.xlsx");
            }
        }

        private void deletelistitmsbycondition()
        {
            string siteurl = "https://appliedapps.amat.com/sites/SSPWAS";

            string listName = "Requests";

            List<string> BCSNumbers = new List<string>();
            BCSNumbers.Add("Test- email notifications with task request for reviewers");
            //BCSNumbers.Add("X054244");
            //BCSNumbers.Add("X069129");
            //BCSNumbers.Add("X056545");
            //BCSNumbers.Add("X056565");
            //BCSNumbers.Add("X052874");
            //BCSNumbers.Add("X073163");
            //BCSNumbers.Add("X064677");
            //BCSNumbers.Add("X076546");
            //BCSNumbers.Add("X083678");
            //BCSNumbers.Add("X083890");
            //BCSNumbers.Add("X056398");
            //BCSNumbers.Add("X084987");
            //BCSNumbers.Add("X066458");
            //BCSNumbers.Add("X085672");
            //BCSNumbers.Add("X086804");
            //BCSNumbers.Add("X087781");
            //BCSNumbers.Add("X087782");
            //BCSNumbers.Add("X088345");
            //BCSNumbers.Add("X089786");
            //BCSNumbers.Add("X090014");
            //BCSNumbers.Add("X082899");
            //BCSNumbers.Add("X069402");
            //BCSNumbers.Add("X091424");
            //BCSNumbers.Add("X092547");
            //BCSNumbers.Add("X092900");
            //BCSNumbers.Add("X094117");
            //BCSNumbers.Add("X094105");
            //BCSNumbers.Add("X094339");
            //BCSNumbers.Add("X094395");
            //BCSNumbers.Add("X096303");
            //BCSNumbers.Add("X096779");
            //BCSNumbers.Add("X096608");
            //BCSNumbers.Add("X097870");
            //BCSNumbers.Add("X098129");
            //BCSNumbers.Add("X077683");
            //BCSNumbers.Add("X099554");
            //BCSNumbers.Add("X096304");
            //BCSNumbers.Add("X0100200");
            //BCSNumbers.Add("x091186");
            //BCSNumbers.Add("X072470");
            //BCSNumbers.Add("X0101891");
            //BCSNumbers.Add("X0100103");
            //BCSNumbers.Add("x0103010");
            //BCSNumbers.Add("X0103822");
            //BCSNumbers.Add("102776");
            //BCSNumbers.Add("102907");
            //BCSNumbers.Add("102968");
            //BCSNumbers.Add("103858");
            //BCSNumbers.Add("105772");
            //BCSNumbers.Add("108025");
            //BCSNumbers.Add("110237");
            //BCSNumbers.Add("70862");
            //BCSNumbers.Add("119865");

            ClientContext clientContext = new ClientContext(siteurl);

            List oList = clientContext.Web.Lists.GetByTitle(listName);    
           

            foreach (string employeeid in BCSNumbers)

            {
                try
                {
                    CamlQuery camlQuery = new CamlQuery();
                    camlQuery.ViewXml = "<View><Query><Where><Eq><FieldRef Name='Title'/>" +
                        "<Value Type='Text'>" + employeeid + "</Value></Eq></Where></Query></View>";


                    ListItemCollection collListItem = oList.GetItems(camlQuery);

                    clientContext.Load(collListItem);

                    clientContext.ExecuteQuery();
                    collListItem[0].DeleteObject();

                    clientContext.ExecuteQuery();
                }
                catch(Exception ex)
                {

                }
            }

    

        }

        private void insertistems()
        {
            string siteurl = "http://amindapps/sites/elms/";

            string listName = "LeaveCreditExclusionList";

    
            

            List<string> BCSNumbers = new List<string>();

            BCSNumbers.Add("108438");
            BCSNumbers.Add("109468");
            BCSNumbers.Add("116320");
            BCSNumbers.Add("121183");
            BCSNumbers.Add("121213");
            BCSNumbers.Add("155703");
            BCSNumbers.Add("156405");
            BCSNumbers.Add("156767");
            BCSNumbers.Add("157161");
            BCSNumbers.Add("157344");
            BCSNumbers.Add("157357");
            BCSNumbers.Add("157584");
            BCSNumbers.Add("157784");
            BCSNumbers.Add("158036");
            BCSNumbers.Add("158388");
            BCSNumbers.Add("158678");
            BCSNumbers.Add("158690");
            BCSNumbers.Add("158691");
            BCSNumbers.Add("158698");
            BCSNumbers.Add("158816");
            BCSNumbers.Add("158825");
            BCSNumbers.Add("158845");
            BCSNumbers.Add("158856");
            BCSNumbers.Add("158870");
            BCSNumbers.Add("158950");
            BCSNumbers.Add("158965");
            BCSNumbers.Add("159169");
            BCSNumbers.Add("159188");
            BCSNumbers.Add("159212");
            BCSNumbers.Add("159218");
            BCSNumbers.Add("159358");
            BCSNumbers.Add("159604");
            BCSNumbers.Add("159833");
            BCSNumbers.Add("159885");
            BCSNumbers.Add("160489");
            BCSNumbers.Add("160610");
            BCSNumbers.Add("160611");
            BCSNumbers.Add("160612");
            BCSNumbers.Add("160614");
            BCSNumbers.Add("160618");
            BCSNumbers.Add("160619");
            BCSNumbers.Add("160620");
            BCSNumbers.Add("160621");
            BCSNumbers.Add("160635");
            BCSNumbers.Add("160638");
            BCSNumbers.Add("160639");
            BCSNumbers.Add("160642");
            BCSNumbers.Add("160746");
            BCSNumbers.Add("160770");
            BCSNumbers.Add("160833");
            BCSNumbers.Add("160834");
            BCSNumbers.Add("160996");
            BCSNumbers.Add("161005");
            BCSNumbers.Add("161006");
            BCSNumbers.Add("161008");
            BCSNumbers.Add("161016");
            BCSNumbers.Add("161031");
            BCSNumbers.Add("161043");
            BCSNumbers.Add("161057");
            BCSNumbers.Add("161058");
            BCSNumbers.Add("161064");
            BCSNumbers.Add("161073");
            BCSNumbers.Add("161131");
            BCSNumbers.Add("161138");
            BCSNumbers.Add("161165");
            BCSNumbers.Add("161177");
            BCSNumbers.Add("161208");
            BCSNumbers.Add("161209");
            BCSNumbers.Add("161210");
            BCSNumbers.Add("161245");
            BCSNumbers.Add("161252");
            BCSNumbers.Add("161311");
            BCSNumbers.Add("161361");
            BCSNumbers.Add("161362");
            BCSNumbers.Add("161369");
            BCSNumbers.Add("161376");
            BCSNumbers.Add("161377");
            BCSNumbers.Add("161413");
            BCSNumbers.Add("161448");
            BCSNumbers.Add("161466");
            BCSNumbers.Add("161469");
            BCSNumbers.Add("161501");
            BCSNumbers.Add("161589");
            BCSNumbers.Add("161630");
            BCSNumbers.Add("161791");
            BCSNumbers.Add("162160");
            BCSNumbers.Add("162162");
            BCSNumbers.Add("162165");
            BCSNumbers.Add("162228");
            BCSNumbers.Add("162260");
            BCSNumbers.Add("162410");
            BCSNumbers.Add("162557");
            BCSNumbers.Add("162808");
            BCSNumbers.Add("162829");
            BCSNumbers.Add("162945");
            BCSNumbers.Add("163011");
            BCSNumbers.Add("163127");
            BCSNumbers.Add("163133");
            BCSNumbers.Add("163263");
            BCSNumbers.Add("163377");
            BCSNumbers.Add("163840");
            BCSNumbers.Add("163972");
            BCSNumbers.Add("164031");
            BCSNumbers.Add("164361");
            BCSNumbers.Add("164534");


            ClientContext clientContext = new ClientContext(siteurl);

            List oList = clientContext.Web.Lists.GetByTitle(listName);
           // ListItem oListItem1 = oList.GetItemById(13579); //13579 //13585

            //// oListItem1["SAF_x0020_Document_x0020_Status"] = "Approved";
            //oListItem1["ClosedOn"] = Convert.ToDateTime("07/11/2019");
            //oListItem1.Update();
            //clientContext.ExecuteQuery();
            // oListItem1 = oList.GetItemById(45070);
            //oListItem1["WorkflowOutcome"] = "Approved";
            //oListItem1.Update();
            //clientContext.ExecuteQuery();
          
            foreach (string title in BCSNumbers)

            {
                try
                {
                    ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                    ListItem oListItem = oList.AddItem(itemCreateInfo);
                    oListItem["Title"] = title;
                    oListItem.Update();

                    clientContext.ExecuteQuery();
                }
                catch (Exception ex)
                {

                }
            }
        }

    }
}
  
