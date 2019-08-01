using SAP.Middleware.Connector;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SAPconnector
{
    class Program
    {
        static void Main(string[] args)
        {
            GetSAPCostCenter("0000380186","Q08");
        }

        public static DataTable GetSAPCostCenter(string CostCenter, string DestinationName)
        {

            RfcDestination dest;
            IRfcFunction bapiFunction;
            dest = RfcDestinationManager.GetDestination(DestinationName);
            RfcRepository rfcRepo = dest.Repository;
            bapiFunction = rfcRepo.CreateFunction("Z_AW_RFC_READ_TABLE");
            DataTable dtCostCenterPartial1;
            IRfcTable tblCostCenter;

            bapiFunction.SetValue("QUERY_TABLE", "CSKS");
            bapiFunction.SetValue("DELIMITER", "~");
            bapiFunction.SetValue("NO_DATA", "");
            bapiFunction.SetValue("ROWSKIPS", "0");
            //bapiFunction.SetValue("ROWCOUNT", "");

            // Parameter table FIELDS contains the columns you want to receive
            // here we query 2 fields, KUNNR and NAME1
            IRfcTable fieldsTable = bapiFunction.GetTable("FIELDS");
            fieldsTable.Append();
            fieldsTable.SetValue("FIELDNAME", "KOKRS");
            fieldsTable.Append();
            fieldsTable.SetValue("FIELDNAME", "KOSTL");
            fieldsTable.Append();
            fieldsTable.SetValue("FIELDNAME", "DATAB");
            fieldsTable.Append();
            fieldsTable.SetValue("FIELDNAME", "BKZKP");
            fieldsTable.Append();
            fieldsTable.SetValue("FIELDNAME", "BUKRS");
            fieldsTable.Append();
            fieldsTable.SetValue("FIELDNAME", "KOSAR");
            fieldsTable.Append();
            fieldsTable.SetValue("FIELDNAME", "KHINR");


            IRfcTable optsTable = bapiFunction.GetTable("OPTIONS");
            optsTable.Append();
            optsTable.SetValue("TEXT", "KOSTL = '" + CostCenter + "'");

            bapiFunction.Invoke(dest);

            tblCostCenter = bapiFunction["DATA"].GetTable();
            IRfcTable dataTable = bapiFunction.GetTable("DATA");
            dtCostCenterPartial1 = new DataTable();
            dtCostCenterPartial1.TableName = "CSKS";
            dtCostCenterPartial1.Columns.Add("Controlling_Area");
            dtCostCenterPartial1.Columns.Add("Cost_Center");
            dtCostCenterPartial1.Columns.Add("Valid_From");
            dtCostCenterPartial1.Columns.Add("Lock_Act_PCosts");
            dtCostCenterPartial1.Columns.Add("Company_Code");
            dtCostCenterPartial1.Columns.Add("CCtr_Category");
            dtCostCenterPartial1.Columns.Add("Std_Hierarchy");
            foreach (var dataRow in dataTable)
            {
                string data = Convert.ToString(dataRow.GetValue("WA"));
                DataRow dr = dtCostCenterPartial1.NewRow();
                string[] columns = data.Split('~');
                if (columns.Count() >= 0)
                    dr["Controlling_Area"] = columns[0];
                if (columns.Count() >= 1)
                    dr["Cost_Center"] = columns[1];
                if (columns.Count() >= 2)
                    dr["Valid_From"] = columns[2];
                if (columns.Count() >= 3)
                    dr["Lock_Act_PCosts"] = columns[3];
                if (columns.Count() >= 4)
                    dr["Company_Code"] = columns[4];
                if (columns.Count() >= 5)
                    dr["CCtr_Category"] = columns[5];
                if (columns.Count() >= 6)
                    dr["Std_Hierarchy"] = columns[6];
                dtCostCenterPartial1.Rows.Add(dr);
            }
            return dtCostCenterPartial1;
        }
    }
}
