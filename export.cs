using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using iTextSharp.text;

using iTextSharp.text.pdf;

using iTextSharp.text.html;

using iTextSharp.text.html.simpleparser;
using System.IO;

namespace cooperativesocietysoftware
{
    public partial class contributiondetails : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            loginname.Text = Session["user"].ToString();
        }

        protected void Button1_Click(object sender, EventArgs e)

        {
            GridView1.DataSourceID = "SqlDataSource4";
            exportwrd.Visible = true;
            exportexcel.Visible = true;
            exportpdf.Visible = true;
        }



        private decimal Totalcontribution = (decimal)0.0;
        private decimal Totaldeduction = (decimal)0.0;
        private decimal Totalbalance = (decimal)0.0;
         protected void GridView1_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            // check row type
            if (e.Row.RowType == DataControlRowType.DataRow)
            { 
                // if row type is DataRow, add ProductSales value to TotalSales
                Totalcontribution += Convert.ToDecimal(DataBinder.Eval(e.Row.DataItem, "creditamount"));
                Totaldeduction += Convert.ToDecimal(DataBinder.Eval(e.Row.DataItem, "debitamount"));
                Totalbalance += Convert.ToDecimal(DataBinder.Eval(e.Row.DataItem, "runningbalance"));
                e.Row.Cells[4].Text = Totalbalance.ToString("N2");

            }  
        
        
            else if (e.Row.RowType == DataControlRowType.Footer)
        {
                // If row type is footer, show calculated total value
                // Since this example uses sales in dollars, I formatted output as currency
                e.Row.Cells[1].Text = "Total";
                 e.Row.Cells[1].Font.Bold = true;
                e.Row.Cells[2].Text = Totaldeduction.ToString("N2");
                e.Row.Cells[3].Text = Totalcontribution.ToString("N2");
                e.Row.Cells[4].Text = Totalbalance.ToString("N2");
                e.Row.Cells[2].Font.Bold = true;
                e.Row.Cells[3].Font.Bold = true;
                e.Row.Cells[4].Font.Bold = true;
                
            }   

        }

        protected void Button2_Click(object sender, EventArgs e)
        {
            string filename = "Test.xls";
            System.IO.StringWriter tw = new System.IO.StringWriter();
            System.Web.UI.HtmlTextWriter hw = new System.Web.UI.HtmlTextWriter(tw);

            //Get the H`enter code here`TML for the control.
            GridView1.RenderControl(hw);
            //Write the HTML back to the browser.
            Response.ContentType = "application/vnd.ms-excel";
            Response.AppendHeader("Content-Disposition", "attachment; filename=" + filename + "");

            Response.Write(tw.ToString());
        }

        public override void VerifyRenderingInServerForm(Control control)

        {

            /* Verifies that the control is rendered */

        }
        protected void exportwrd_Click(object sender, EventArgs e)
        {
            Response.Clear();

            Response.Buffer = true;

            Response.AddHeader("content-disposition",

            "attachment;filename=GridViewExport.doc");

            Response.Charset = "";

            Response.ContentType = "application/vnd.ms-word ";

            StringWriter sw = new StringWriter();

            HtmlTextWriter hw = new HtmlTextWriter(sw);

            GridView1.AllowPaging = false;

            GridView1.DataBind();

            GridView1.RenderControl(hw);

            Response.Output.Write(sw.ToString());

            Response.Flush();

            Response.End();
        }

        protected void exportexcel_Click(object sender, EventArgs e)
        {
            Response.Clear();

            Response.Buffer = true;



            Response.AddHeader("content-disposition",

            "attachment;filename=GridViewExport.xls");

            Response.Charset = "";

            Response.ContentType = "application/vnd.ms-excel";

            StringWriter sw = new StringWriter();

            HtmlTextWriter hw = new HtmlTextWriter(sw);



            GridView1.AllowPaging = false;

            GridView1.DataBind();



            //Change the Header Row back to white color

            GridView1.HeaderRow.Style.Add("background-color", "#FFFFFF");



            //Apply style to Individual Cells

            GridView1.HeaderRow.Cells[0].Style.Add("background-color", "green");

            GridView1.HeaderRow.Cells[1].Style.Add("background-color", "green");

            GridView1.HeaderRow.Cells[2].Style.Add("background-color", "green");

            GridView1.HeaderRow.Cells[3].Style.Add("background-color", "green");



            for (int i = 0; i < GridView1.Rows.Count; i++)

            {

                GridViewRow row = GridView1.Rows[i];



                //Change Color back to white

                row.BackColor = System.Drawing.Color.White;



                //Apply text style to each Row

                row.Attributes.Add("class", "textmode");



                //Apply style to Individual Cells of Alternating Row

                if (i % 2 != 0)

                {

                    row.Cells[0].Style.Add("background-color", "#C2D69B");

                    row.Cells[1].Style.Add("background-color", "#C2D69B");

                    row.Cells[2].Style.Add("background-color", "#C2D69B");

                    row.Cells[3].Style.Add("background-color", "#C2D69B");

                }

            }

            GridView1.RenderControl(hw);



            //style to format numbers to string

            string style = @"<style> .textmode { mso-number-format:\@; } </style>";

            Response.Write(style);

            Response.Output.Write(sw.ToString());

            Response.Flush();

            Response.End();
        }

        protected void exportpdf_Click(object sender, EventArgs e)
        {
            Response.ContentType = "application/pdf";

            Response.AddHeader("content-disposition",

             "attachment;filename=GridViewExport.pdf");

            Response.Cache.SetCacheability(HttpCacheability.NoCache);

            StringWriter sw = new StringWriter();

            HtmlTextWriter hw = new HtmlTextWriter(sw);

            GridView1.AllowPaging = false;

            GridView1.DataBind();

            GridView1.RenderControl(hw);

            StringReader sr = new StringReader(sw.ToString());

            Document pdfDoc = new Document(PageSize.A4, 10f, 10f, 10f, 0f);

            HTMLWorker htmlparser = new HTMLWorker(pdfDoc);

            PdfWriter.GetInstance(pdfDoc, Response.OutputStream);

            pdfDoc.Open();

            htmlparser.Parse(sr);

            pdfDoc.Close();

            Response.Write(pdfDoc);

            Response.End();
        }
    }
}
