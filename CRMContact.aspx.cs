using ExcelDataReader;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Hosting;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace D365ContactExcelWep
{
    public partial class CRMContact : System.Web.UI.Page
    {
        string connectionString = "AuthType=Office365;Url=https://icoskun.crm11.dynamics.com/;Username=CumhurCoskun@RuzicoLimited671.onmicrosoft.com;Password=Telefon1;";

        protected void Page_Load(object sender, EventArgs e)
        {

            if (!IsPostBack)
            {
                BindContactGridView();
            }
        }

        private void BindContactGridView()
        {
            CrmServiceClient crmServiceClient = new CrmServiceClient(connectionString);

            if (!crmServiceClient.IsReady)
            {
                resultLabel.Text = "Failed to establish CRM connection.";
                return;
            }

            QueryExpression query = new QueryExpression("contact");
            query.ColumnSet = new ColumnSet("firstname", "lastname", "emailaddress1", "createdon");
            query.Orders.Add(new OrderExpression("createdon", OrderType.Descending));

            EntityCollection contacts = crmServiceClient.RetrieveMultiple(query);
            if (contacts.Entities.Count >= 1)
            {
                DataSet ds = new DataSet();
                DataTable dt = new DataTable();
                dt.Columns.Add("firstname");
                dt.Columns.Add("lastname");
                dt.Columns.Add("emailaddress1");
                dt.Columns.Add("createdon");

                foreach (var ContactResultUsingCallerId in contacts.Entities)
                {
                    DataRow dr = dt.NewRow();

                    dr["firstname"] = ContactResultUsingCallerId.GetAttributeValue<string>("firstname");
                    dr["lastname"] = ContactResultUsingCallerId.GetAttributeValue<string>("lastname");
                    dr["emailaddress1"] = ContactResultUsingCallerId.GetAttributeValue<string>("emailaddress1"); 
                    dr["createdon"] = ContactResultUsingCallerId.GetAttributeValue<DateTime>("createdon").ToString();

                    dt.Rows.Add(dr);
                }

                ds.Tables.Add(dt);

                // You are now ready to bind your DataSet to your GridView
                contactGridView.DataSource = ds;
                contactGridView.DataBind();
            }

        }


        protected void UploadButton_Click(object sender, EventArgs e)
        {
            if (fileUpload.HasFile)
            {
                string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(fileUpload.PostedFile.FileName);
                string extension = Path.GetExtension(fileUpload.PostedFile.FileName);
                string newFileName = fileNameWithoutExtension + "_" + Guid.NewGuid().ToString() + extension;

                // Save the uploaded file
                string excelFilePath = HostingEnvironment.MapPath("~/App_Data/" + newFileName);
                fileUpload.SaveAs(excelFilePath);

                // Establish a connection to Dynamics CRM
                CrmServiceClient crmServiceClient = new CrmServiceClient(connectionString);

                if (!crmServiceClient.IsReady)
                {
                    resultLabel.Text = "Failed to establish CRM connection.";
                    return;
                }

                // Read Contact data from the Excel file and create records in Dynamics CRM
                List<Entity> contacts = ReadContactsFromExcel(excelFilePath);

                foreach (Entity contact in contacts)
                {
                    Guid contactId = crmServiceClient.Create(contact);
                    resultLabel.Text += "New contact created. Contact ID: " + contactId + "<br/>";
                }

                // Refresh the GridView with the updated Contacts
                BindContactGridView();
            }
            else
            {
                resultLabel.Text = "No file selected.";
            }
        }

        private List<Entity> ReadContactsFromExcel(string filePath)
        {
            List<Entity> contacts = new List<Entity>();

            using (var stream = System.IO.File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    DataSet result = reader.AsDataSet();
                    DataTable table = result.Tables[0];

                    foreach (DataRow row in table.Rows)
                    {
                        // check if the row is table header
                        if (row != table.Rows[0])
                        {
                            Entity contact = new Entity("contact");
                            contact["firstname"] = row.ItemArray[0].ToString();
                            contact["lastname"] = row.ItemArray[1].ToString();
                            contact["emailaddress1"] = row.ItemArray[2].ToString();
                            contacts.Add(contact);
                        }

                    }
                }
            }

            return contacts;
        }

    }
}