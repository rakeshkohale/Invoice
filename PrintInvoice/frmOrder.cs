using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.OleDb;

namespace PrintInvoice
{
	/// <summary>
	/// Summary description for frmOrder.
	/// </summary>
	public class frmOrder : System.Windows.Forms.Form
	{
		private System.Windows.Forms.DataGrid ordGrid;
		private System.Windows.Forms.Button btnPrint;
		private System.Windows.Forms.Button btnPreview;
		private System.Windows.Forms.Button btnClose;
		private System.Windows.Forms.Button btnDialog;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label lblDate;
		private System.Windows.Forms.Label lblCustomer;
		private System.Windows.Forms.Label lblSeller;
		private System.Windows.Forms.Label lblCity;
		private System.Windows.Forms.Label lblID;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label lblSubTotal;
		private System.Windows.Forms.Label lblInvoiceTotal;
		private System.Windows.Forms.Label lblFreight;

		// for PrintDialog, PrintPreviewDialog and PrintDocument:
		private System.Windows.Forms.PrintDialog prnDialog;
		private System.Windows.Forms.PrintPreviewDialog prnPreview;
		private System.Drawing.Printing.PrintDocument prnDocument;
		private System.ComponentModel.Container components = null;
		
		// for Invoice Head:
		private string InvTitle;
		private string InvSubTitle1;
		private string InvSubTitle2;
		private string InvSubTitle3;
		private string InvImage;

		// for Database:
		private OleDbConnection cnn;
		private OleDbCommand cmd;
		private OleDbDataReader rdrInvoice;
		private string strCon;
		private string InvSql;

		// for Report:
		private int CurrentY;
		private int CurrentX;
		private int leftMargin;
		private int rightMargin;
		private int topMargin;
		private int bottomMargin;
		private int InvoiceWidth;
		private int InvoiceHeight;
		private string CustomerName;
		private string CustomerCity;
		private string SellerName;
		private string SaleID;
		private string SaleDate;
		private decimal SaleFreight;
		private decimal SubTotal;
		private decimal InvoiceTotal;
		private bool ReadInvoice;
		private int AmountPosition;
		
		// Font and Color:------------------
		// Title Font
		private Font InvTitleFont = new Font("Arial", 24, FontStyle.Regular);
		// Title Font height
		private int InvTitleHeight;
		// SubTitle Font
		private Font InvSubTitleFont = new Font("Arial", 14, FontStyle.Regular);
		// SubTitle Font height
		private int InvSubTitleHeight;
		// Invoice Font
		private Font InvoiceFont = new Font("Arial", 12, FontStyle.Regular);
		// Invoice Font height
		private int InvoiceFontHeight;
		// Blue Color
		private SolidBrush BlueBrush = new SolidBrush(Color.Blue);
		// Red Color
		private SolidBrush RedBrush = new SolidBrush(Color.Red);
		// Black Color
		private SolidBrush BlackBrush = new SolidBrush(Color.Black);
        private string testId;

		public frmOrder()
		{
			InitializeComponent();

			this.prnDialog = new System.Windows.Forms.PrintDialog();
			this.prnPreview = new System.Windows.Forms.PrintPreviewDialog();
			this.prnDocument = new System.Drawing.Printing.PrintDocument();
			// The Event of 'PrintPage'
			prnDocument.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(prnDocument_PrintPage);
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if(components != null)
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.ordGrid = new System.Windows.Forms.DataGrid();
			this.btnPrint = new System.Windows.Forms.Button();
			this.btnPreview = new System.Windows.Forms.Button();
			this.btnClose = new System.Windows.Forms.Button();
			this.btnDialog = new System.Windows.Forms.Button();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.lblDate = new System.Windows.Forms.Label();
			this.lblCustomer = new System.Windows.Forms.Label();
			this.lblSeller = new System.Windows.Forms.Label();
			this.lblCity = new System.Windows.Forms.Label();
			this.lblID = new System.Windows.Forms.Label();
			this.label9 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.lblSubTotal = new System.Windows.Forms.Label();
			this.lblInvoiceTotal = new System.Windows.Forms.Label();
			this.lblFreight = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.ordGrid)).BeginInit();
			this.SuspendLayout();
			// 
			// ordGrid
			// 
			this.ordGrid.DataMember = "";
			this.ordGrid.HeaderForeColor = System.Drawing.SystemColors.ControlText;
			this.ordGrid.Location = new System.Drawing.Point(16, 88);
			this.ordGrid.Name = "ordGrid";
			this.ordGrid.Size = new System.Drawing.Size(480, 200);
			this.ordGrid.TabIndex = 37;
			// 
			// btnPrint
			// 
			this.btnPrint.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnPrint.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.btnPrint.Location = new System.Drawing.Point(112, 320);
			this.btnPrint.Name = "btnPrint";
			this.btnPrint.Size = new System.Drawing.Size(88, 20);
			this.btnPrint.TabIndex = 39;
			this.btnPrint.Text = "Print";
			this.btnPrint.Click += new System.EventHandler(this.btnPrint_Click);
			// 
			// btnPreview
			// 
			this.btnPreview.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnPreview.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.btnPreview.Location = new System.Drawing.Point(16, 320);
			this.btnPreview.Name = "btnPreview";
			this.btnPreview.Size = new System.Drawing.Size(88, 20);
			this.btnPreview.TabIndex = 38;
			this.btnPreview.Text = "Print Preview";
			this.btnPreview.Click += new System.EventHandler(this.btnPreview_Click);
			// 
			// btnClose
			// 
			this.btnClose.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnClose.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.btnClose.Location = new System.Drawing.Point(408, 320);
			this.btnClose.Name = "btnClose";
			this.btnClose.Size = new System.Drawing.Size(88, 20);
			this.btnClose.TabIndex = 40;
			this.btnClose.Text = "Close";
			this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
			// 
			// btnDialog
			// 
			this.btnDialog.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnDialog.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.btnDialog.Location = new System.Drawing.Point(312, 320);
			this.btnDialog.Name = "btnDialog";
			this.btnDialog.Size = new System.Drawing.Size(88, 20);
			this.btnDialog.TabIndex = 41;
			this.btnDialog.Text = "Print Dialog";
			this.btnDialog.Click += new System.EventHandler(this.btnDialog_Click);
			// 
			// label1
			// 
			this.label1.ForeColor = System.Drawing.Color.Navy;
			this.label1.Location = new System.Drawing.Point(16, 8);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(96, 20);
			this.label1.TabIndex = 42;
			this.label1.Text = "Company Name:";
			// 
			// label2
			// 
			this.label2.ForeColor = System.Drawing.Color.Navy;
			this.label2.Location = new System.Drawing.Point(336, 8);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(32, 20);
			this.label2.TabIndex = 43;
			this.label2.Text = "City:";
			// 
			// label3
			// 
			this.label3.ForeColor = System.Drawing.Color.Navy;
			this.label3.Location = new System.Drawing.Point(16, 56);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(56, 20);
			this.label3.TabIndex = 44;
			this.label3.Text = "Order ID:";
			// 
			// label4
			// 
			this.label4.ForeColor = System.Drawing.Color.Navy;
			this.label4.Location = new System.Drawing.Point(16, 32);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(96, 20);
			this.label4.TabIndex = 45;
			this.label4.Text = "Salesperson:";
			// 
			// label5
			// 
			this.label5.ForeColor = System.Drawing.Color.Navy;
			this.label5.Location = new System.Drawing.Point(336, 56);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(64, 20);
			this.label5.TabIndex = 46;
			this.label5.Text = "Order Date:";
			// 
			// lblDate
			// 
			this.lblDate.BackColor = System.Drawing.Color.White;
			this.lblDate.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.lblDate.Location = new System.Drawing.Point(400, 56);
			this.lblDate.Name = "lblDate";
			this.lblDate.Size = new System.Drawing.Size(96, 20);
			this.lblDate.TabIndex = 53;
			this.lblDate.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblCustomer
			// 
			this.lblCustomer.BackColor = System.Drawing.Color.White;
			this.lblCustomer.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.lblCustomer.Location = new System.Drawing.Point(112, 8);
			this.lblCustomer.Name = "lblCustomer";
			this.lblCustomer.Size = new System.Drawing.Size(216, 20);
			this.lblCustomer.TabIndex = 54;
			this.lblCustomer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblSeller
			// 
			this.lblSeller.BackColor = System.Drawing.Color.White;
			this.lblSeller.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.lblSeller.Location = new System.Drawing.Point(112, 32);
			this.lblSeller.Name = "lblSeller";
			this.lblSeller.Size = new System.Drawing.Size(216, 20);
			this.lblSeller.TabIndex = 55;
			this.lblSeller.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblCity
			// 
			this.lblCity.BackColor = System.Drawing.Color.White;
			this.lblCity.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.lblCity.Location = new System.Drawing.Point(368, 8);
			this.lblCity.Name = "lblCity";
			this.lblCity.Size = new System.Drawing.Size(128, 20);
			this.lblCity.TabIndex = 56;
			this.lblCity.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblID
			// 
			this.lblID.BackColor = System.Drawing.Color.White;
			this.lblID.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.lblID.Location = new System.Drawing.Point(72, 56);
			this.lblID.Name = "lblID";
			this.lblID.Size = new System.Drawing.Size(96, 20);
			this.lblID.TabIndex = 57;
			this.lblID.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label9
			// 
			this.label9.ForeColor = System.Drawing.Color.Navy;
			this.label9.Location = new System.Drawing.Point(368, 292);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(40, 20);
			this.label9.TabIndex = 64;
			this.label9.Text = "Total:";
			// 
			// label8
			// 
			this.label8.ForeColor = System.Drawing.Color.Navy;
			this.label8.Location = new System.Drawing.Point(192, 292);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(48, 20);
			this.label8.TabIndex = 63;
			this.label8.Text = "Freight:";
			// 
			// label7
			// 
			this.label7.ForeColor = System.Drawing.Color.Navy;
			this.label7.Location = new System.Drawing.Point(16, 292);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(56, 20);
			this.label7.TabIndex = 62;
			this.label7.Text = "Subtotal:";
			// 
			// lblSubTotal
			// 
			this.lblSubTotal.BackColor = System.Drawing.Color.White;
			this.lblSubTotal.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.lblSubTotal.Location = new System.Drawing.Point(72, 292);
			this.lblSubTotal.Name = "lblSubTotal";
			this.lblSubTotal.Size = new System.Drawing.Size(88, 20);
			this.lblSubTotal.TabIndex = 65;
			this.lblSubTotal.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// lblInvoiceTotal
			// 
			this.lblInvoiceTotal.BackColor = System.Drawing.Color.White;
			this.lblInvoiceTotal.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.lblInvoiceTotal.Location = new System.Drawing.Point(406, 292);
			this.lblInvoiceTotal.Name = "lblInvoiceTotal";
			this.lblInvoiceTotal.Size = new System.Drawing.Size(88, 20);
			this.lblInvoiceTotal.TabIndex = 66;
			this.lblInvoiceTotal.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// lblFreight
			// 
			this.lblFreight.BackColor = System.Drawing.Color.White;
			this.lblFreight.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.lblFreight.Location = new System.Drawing.Point(240, 292);
			this.lblFreight.Name = "lblFreight";
			this.lblFreight.Size = new System.Drawing.Size(88, 20);
			this.lblFreight.TabIndex = 67;
			this.lblFreight.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// frmOrder
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(512, 351);
			this.Controls.Add(this.lblFreight);
			this.Controls.Add(this.lblInvoiceTotal);
			this.Controls.Add(this.lblSubTotal);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.lblID);
			this.Controls.Add(this.lblCity);
			this.Controls.Add(this.lblSeller);
			this.Controls.Add(this.lblCustomer);
			this.Controls.Add(this.lblDate);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.btnDialog);
			this.Controls.Add(this.btnClose);
			this.Controls.Add(this.btnPrint);
			this.Controls.Add(this.btnPreview);
			this.Controls.Add(this.ordGrid);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.MaximizeBox = false;
			this.Name = "frmOrder";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Invoice";
			this.Load += new System.EventHandler(this.frmOrder_Load);
			((System.ComponentModel.ISupportInitialize)(this.ordGrid)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
		
		private void LoadOrder()
		{
			int intOrder = int.Parse(frmInput.InvoiceOrder);

			// following lines to connect with access database file 'Northwind.mdb' 
			//string MyDataFile = Application.StartupPath + @"\DataFile\Northwind.mdb";
			//string MyPass = "";
			//strCon = @"provider=microsoft.jet.oledb.4.0;data source=" + 
				//MyDataFile + ";" + "Jet OLEDB:Database Password=" + MyPass + ";";

            // If you are using SQL Server, please replace previous lines with following
            strCon = @"provider=sqloledb;Data Source=uatdb24.kotakseconline.com;Initial Catalog=KSCS_DB;User ID=kscs_appowner;Password=kscs_appowner@123";
            //and replace 'Data Source=PC' with the name of your system */

            try
			{
				// Get Invoice Data:
                InvSql = "SELECT * from mstestdetails where testid = " + intOrder; //testId;
				
				//create an OleDbDataAdapter
				OleDbDataAdapter datAdp = new OleDbDataAdapter(InvSql, strCon);

				//create a command builder
				OleDbCommandBuilder cBuilder = new OleDbCommandBuilder(datAdp);

				//create a DataTable to hold the query results
				DataTable dTable = new DataTable();

				//fill the DataTable
				datAdp.Fill(dTable);

				if (dTable.Rows.Count == 0)
				{
					MessageBox.Show("This Order not found, Please enter another order.");
					this.Close();
				}

				// Create a TableStyle to format Datagrid columns.
				ordGrid.TableStyles.Clear();
				DataGridTableStyle tableStyle = new DataGridTableStyle();

				foreach(DataColumn dc in dTable.Columns)
				{
					DataGridTextBoxColumn txtColumn = new DataGridTextBoxColumn();
					txtColumn.MappingName = dc.ColumnName;
					txtColumn.HeaderText = dc.Caption;
					switch (dc.ColumnName.ToString())
					{
                        case "TestID":   // Product ID 
							txtColumn.HeaderText = "Product ID";
							txtColumn.Width = 60;
							break;
                        case "TestLevelID":   // Product Name 
							txtColumn.HeaderText = "Product Name";
							txtColumn.Width = 110;
							break;
                        case "StartDate":   // Unit Price 
							txtColumn.HeaderText = "Unit Price";
							txtColumn.Format = "0.00";
							txtColumn.Alignment = HorizontalAlignment.Right;
							txtColumn.Width = 60;
							break;
                        case "EndDate":   // Discount 
							txtColumn.HeaderText = "Discount";
							txtColumn.Format = "p"; // Percent
							txtColumn.Alignment = HorizontalAlignment.Right;
							txtColumn.Width = 60;
							break;
                        case "TrStatus":   // Quantity 
							txtColumn.HeaderText = "Quantity";
							txtColumn.Alignment = HorizontalAlignment.Right;
							txtColumn.Width = 50;
							break;
                        case "Duration":   // Extended Price 
							txtColumn.HeaderText = "Extended Price";
							txtColumn.Format = "0.00";
							txtColumn.Alignment = HorizontalAlignment.Right;
							txtColumn.Width = 90;
							break;
					}
					tableStyle.GridColumnStyles.Add(txtColumn);
				}

				tableStyle.MappingName = dTable.TableName;
				ordGrid.TableStyles.Add(tableStyle);
				//set DataSource of DataGrid 
				ordGrid.DataSource = dTable.DefaultView;
			}
			catch(Exception e)
			{ 
				MessageBox.Show(e.ToString());
			}
		}
		
		private void FindOrderData()
		{
			//int intOrder = int.Parse(frmInput.InvoiceOrder);
            int testid = int.Parse(frmInput.InvoiceOrder);

            string InvSql = "SELECT * FROM mstest WHERE testID = " +testid; //intOrder;

			OleDbConnection cnn = new OleDbConnection(strCon);
			OleDbCommand cmdOrder = new OleDbCommand(InvSql, cnn);
			cnn.Open();
			OleDbDataReader rdrOrder = cmdOrder.ExecuteReader();

			// Get CompanyName, City, Salesperson, OrderID, OrderDate and Freight
			rdrOrder.Read();
			CustomerName = rdrOrder["TestID"].ToString();
			CustomerCity = rdrOrder["TestLevelID"].ToString();
			SellerName = rdrOrder["StartDate"].ToString();
			SaleID = rdrOrder["EndDate"].ToString();
            //System.DateTime dtOrder = Convert.ToDateTime(rdrOrder["TrStatus"]);
            //SaleDate = dtOrder.ToShortDateString();
            //SaleFreight = Convert.ToDecimal(rdrOrder["Freight"]);
			// Get invoice total
			GetInvoiceTotal();

			rdrOrder.Close();
			cnn.Close();
		}

		private void ReadInvoiceHead()
		{
			//Titles and Image of invoice:
			InvTitle = "iStart Incorporatiom";
			InvSubTitle1 = "23 Nerul";
			InvSubTitle2 = "Navi Mumbai, India";
			InvSubTitle3 = "Phone 1111222233";
			InvImage = Application.StartupPath + @"\Images\" + "kscs.jpg";
		}

		private void GetInvoiceTotal()
		{
			SubTotal = 0;

			cnn = new OleDbConnection(strCon);
			cmd = new OleDbCommand(InvSql, cnn);
			cnn.Open();
			rdrInvoice = cmd.ExecuteReader();

			while (rdrInvoice.Read()) 
			{
                SubTotal = 0;// SubTotal + Convert.ToDecimal(rdrInvoice["ExtendedPrice"]);
			}
			
			rdrInvoice.Close();
			cnn.Close();

			// Get Total
			InvoiceTotal = SubTotal + SaleFreight;
			// Set Total
			lblSubTotal.Text = SubTotal.ToString();
			lblFreight.Text = SaleFreight.ToString();
			lblInvoiceTotal.Text = InvoiceTotal.ToString();
		}

		private void ReadInvoiceData()
		{
			cnn.Open();
			rdrInvoice = cmd.ExecuteReader();
			rdrInvoice.Read();
		}

		private void SetInvoiceHead(Graphics g)
		{
			ReadInvoiceHead();

			CurrentY = topMargin;
			CurrentX = leftMargin;
			int ImageHeight = 0;

			// Draw Invoice image:
			if (System.IO.File.Exists(InvImage))
			{
				Bitmap oInvImage = new Bitmap(InvImage);
				// Set Image Left to center Image:
				int xImage = CurrentX + (InvoiceWidth - oInvImage.Width)/2;
				ImageHeight = oInvImage.Height; // Get Image Height
				g.DrawImage(oInvImage, xImage, CurrentY);
			}

			InvTitleHeight = (int)(InvTitleFont.GetHeight(g));
			InvSubTitleHeight = (int)(InvSubTitleFont.GetHeight(g));

			// Get Titles Length:
			int lenInvTitle = (int)g.MeasureString(InvTitle, InvTitleFont).Width;
			int lenInvSubTitle1 = (int)g.MeasureString(InvSubTitle1, InvSubTitleFont).Width;
			int lenInvSubTitle2 = (int)g.MeasureString(InvSubTitle2, InvSubTitleFont).Width;
			int lenInvSubTitle3 = (int)g.MeasureString(InvSubTitle3, InvSubTitleFont).Width;
			// Set Titles Left:
			int xInvTitle = CurrentX + (InvoiceWidth - lenInvTitle)/2;
			int xInvSubTitle1 = CurrentX + (InvoiceWidth - lenInvSubTitle1)/2;
			int xInvSubTitle2 = CurrentX + (InvoiceWidth - lenInvSubTitle2)/2;
			int xInvSubTitle3 = CurrentX + (InvoiceWidth - lenInvSubTitle3)/2;

			// Draw Invoice Head:
			if(InvTitle != "")
			{
				CurrentY = CurrentY + ImageHeight;
				g.DrawString(InvTitle, InvTitleFont, BlueBrush, xInvTitle, CurrentY);
			}
			if(InvSubTitle1 != "")
			{
				CurrentY = CurrentY + InvTitleHeight;
				g.DrawString(InvSubTitle1, InvSubTitleFont, BlueBrush, xInvSubTitle1, CurrentY);
			}
			if(InvSubTitle2 != "")
			{
				CurrentY = CurrentY + InvSubTitleHeight;
				g.DrawString(InvSubTitle2, InvSubTitleFont, BlueBrush, xInvSubTitle2, CurrentY);
			}
			if(InvSubTitle3 != "")
			{
				CurrentY = CurrentY + InvSubTitleHeight;
				g.DrawString(InvSubTitle3, InvSubTitleFont, BlueBrush, xInvSubTitle3, CurrentY);
			}

			// Draw line:
			CurrentY = CurrentY + InvSubTitleHeight + 8;
			g.DrawLine(new Pen(Brushes.Black, 2), CurrentX, CurrentY, rightMargin, CurrentY);
		}

		private void SetOrderData(Graphics g)
		{// Set Company Name, City, Salesperson, Order ID and Order Date
			string FieldValue = "";
			InvoiceFontHeight = (int)(InvoiceFont.GetHeight(g));
			// Set Company Name:
			CurrentX = leftMargin;
			CurrentY = CurrentY + 8;
			FieldValue = "Company Name: " + CustomerName;
			g.DrawString(FieldValue, InvoiceFont, BlackBrush, CurrentX, CurrentY);
			// Set City:
			CurrentX = CurrentX + (int)g.MeasureString(FieldValue, InvoiceFont).Width + 16;
			FieldValue = "City: " + CustomerCity;
			g.DrawString(FieldValue, InvoiceFont, BlackBrush, CurrentX, CurrentY);
			// Set Salesperson:
			CurrentX = leftMargin;
			CurrentY = CurrentY + InvoiceFontHeight;
			FieldValue = "Salesperson: " + SellerName;
			g.DrawString(FieldValue, InvoiceFont, BlackBrush, CurrentX, CurrentY);
			// Set Order ID:
			CurrentX = leftMargin;
			CurrentY = CurrentY + InvoiceFontHeight;
			FieldValue = "Order ID: " + SaleID;
			g.DrawString(FieldValue, InvoiceFont, BlackBrush, CurrentX, CurrentY);
			// Set Order Date:
			CurrentX = CurrentX + (int)g.MeasureString(FieldValue, InvoiceFont).Width + 16;
			FieldValue = "Order Date: " + SaleDate;
			g.DrawString(FieldValue, InvoiceFont, BlackBrush, CurrentX, CurrentY);
			
			// Draw line:
			CurrentY = CurrentY + InvoiceFontHeight + 8;
			g.DrawLine(new Pen(Brushes.Black), leftMargin, CurrentY, rightMargin, CurrentY);
		}

		private void SetInvoiceData(Graphics g, System.Drawing.Printing.PrintPageEventArgs e)
		{// Set Invoice Table:
			string FieldValue = "";
			int CurrentRecord = 0;
			int RecordsPerPage = 20; // twenty items in a page
			decimal Amount = 0;
			bool StopReading = false;

			// Set Table Head:
			int xProductID = leftMargin;
			CurrentY = CurrentY + InvoiceFontHeight;
			g.DrawString("Product ID", InvoiceFont, BlueBrush, xProductID, CurrentY);

			int xProductName = xProductID + (int)g.MeasureString("Product ID", InvoiceFont).Width + 4;
			g.DrawString("Product Name", InvoiceFont, BlueBrush, xProductName, CurrentY);

			int xUnitPrice = xProductName  + (int)g.MeasureString("Product Name", InvoiceFont).Width + 72;
			g.DrawString("Unit Price", InvoiceFont, BlueBrush, xUnitPrice, CurrentY);

			int xQuantity = xUnitPrice  + (int)g.MeasureString("Unit Price", InvoiceFont).Width + 4;
			g.DrawString("Quantity", InvoiceFont, BlueBrush, xQuantity, CurrentY);

			int xDiscount = xQuantity  + (int)g.MeasureString("Quantity", InvoiceFont).Width + 4;
			g.DrawString("Discount", InvoiceFont, BlueBrush, xDiscount, CurrentY);

			AmountPosition = xDiscount  + (int)g.MeasureString("Discount", InvoiceFont).Width + 4;
			g.DrawString("Extended Price", InvoiceFont, BlueBrush, AmountPosition, CurrentY);

			// Set Invoice Table:
			CurrentY = CurrentY + InvoiceFontHeight + 8;

			while (CurrentRecord < RecordsPerPage)
			{
                FieldValue = rdrInvoice["TestId"].ToString();
				g.DrawString(FieldValue, InvoiceFont, BlackBrush, xProductID, CurrentY);
                FieldValue = rdrInvoice["TestId"].ToString();
				// if Length of (Product Name) > 20, Draw 20 character only
				if (FieldValue.Length > 20)
					FieldValue = FieldValue.Remove(20, FieldValue.Length - 20);
				g.DrawString(FieldValue, InvoiceFont, BlackBrush, xProductName, CurrentY);
                FieldValue = String.Format("{0:0.00}", rdrInvoice["TestId"]); 
				g.DrawString(FieldValue, InvoiceFont, BlackBrush, xUnitPrice, CurrentY);
                FieldValue = rdrInvoice["TestId"].ToString();
				g.DrawString(FieldValue, InvoiceFont, BlackBrush, xQuantity, CurrentY);
                FieldValue = String.Format("{0:0.00%}", rdrInvoice["TestId"]); 
				g.DrawString(FieldValue, InvoiceFont, BlackBrush, xDiscount, CurrentY);

                Amount = Convert.ToDecimal(rdrInvoice["TestId"]);
				// Format Extended Price and Align to Right:
				FieldValue = String.Format("{0:0.00}", Amount); 
				int xAmount = AmountPosition + (int)g.MeasureString("Extended Price", InvoiceFont).Width;
				xAmount = xAmount - (int)g.MeasureString(FieldValue, InvoiceFont).Width;
				g.DrawString(FieldValue, InvoiceFont, BlackBrush, xAmount, CurrentY);
				CurrentY = CurrentY + InvoiceFontHeight;
				
				if (!rdrInvoice.Read())
				{
					StopReading = true;
					break;
				}

				CurrentRecord ++;
			}

			if (CurrentRecord < RecordsPerPage)
				e.HasMorePages = false;
			else
				e.HasMorePages = true;

			if (StopReading)
			{
				rdrInvoice.Close();
				cnn.Close();
				SetInvoiceTotal(g);
			}

			g.Dispose();
		}

		private void SetInvoiceTotal(Graphics g)
		{// Set Invoice Total:
			// Draw line:
			CurrentY = CurrentY + 8;
			g.DrawLine(new Pen(Brushes.Black), leftMargin, CurrentY, rightMargin, CurrentY);
			// Get Right Edge of Invoice:
			int xRightEdg = AmountPosition + (int)g.MeasureString("Extended Price", InvoiceFont).Width;
			
			// Write Sub Total:
			int xSubTotal = AmountPosition  - (int)g.MeasureString("Sub Total", InvoiceFont).Width;
			CurrentY = CurrentY + 8;
			g.DrawString("Sub Total", InvoiceFont, RedBrush, xSubTotal, CurrentY);
			string TotalValue = String.Format("{0:0.00}", SubTotal); 
			int xTotalValue = xRightEdg - (int)g.MeasureString(TotalValue, InvoiceFont).Width;
			g.DrawString(TotalValue, InvoiceFont, BlackBrush, xTotalValue, CurrentY);
			
			// Write Order Freight:
			int xOrderFreight = AmountPosition  - (int)g.MeasureString("Order Freight", InvoiceFont).Width;
			CurrentY = CurrentY + InvoiceFontHeight;
			g.DrawString("Order Freight", InvoiceFont, RedBrush, xOrderFreight, CurrentY);
			string FreightValue = String.Format("{0:0.00}", SaleFreight); 
			int xFreight = xRightEdg - (int)g.MeasureString(FreightValue, InvoiceFont).Width;
			g.DrawString(FreightValue, InvoiceFont, BlackBrush, xFreight, CurrentY);
			
			// Write Invoice Total:
			int xInvoiceTotal = AmountPosition  - (int)g.MeasureString("Invoice Total", InvoiceFont).Width;
			CurrentY = CurrentY + InvoiceFontHeight;
			g.DrawString("Invoice Total", InvoiceFont, RedBrush, xInvoiceTotal, CurrentY);
			string InvoiceValue = String.Format("{0:0.00}", InvoiceTotal); 
			int xInvoiceValue = xRightEdg - (int)g.MeasureString(InvoiceValue, InvoiceFont).Width;
			g.DrawString(InvoiceValue, InvoiceFont, BlackBrush, xInvoiceValue, CurrentY);
		}

		private void DisplayDialog()
		{
			try
			{
				prnDialog.Document = this.prnDocument;
				DialogResult ButtonPressed = prnDialog.ShowDialog();
				// If user Click 'OK', Print Invoice
				if (ButtonPressed == DialogResult.OK)
					prnDocument.Print();
			}
			catch(Exception e) 
			{
				MessageBox.Show(e.ToString());
			}
		}

		private void DisplayInvoice()
		{
			prnPreview.Document = this.prnDocument;

			try
			{
				prnPreview.ShowDialog();
			}
			catch(Exception e) 
			{
				MessageBox.Show(e.ToString());
			}
		}

		private void PrintReport()
		{
			try
			{
				prnDocument.Print();
			}
			catch(Exception e) 
			{
				MessageBox.Show(e.ToString());
			}
		}

		private void frmOrder_Load(object sender, System.EventArgs e)
		{
			ordGrid.CaptionText = "Invoice...";
			LoadOrder();
			FindOrderData();

			lblCustomer.Text = CustomerName;
			lblCity.Text = CustomerCity;
			lblSeller.Text = SellerName;
			lblID.Text = SaleID;
			lblDate.Text = SaleDate;
		}

		// Result of the Event 'PrintPage'
		private void prnDocument_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
		{
			leftMargin = (int)e.MarginBounds.Left;
			rightMargin = (int)e.MarginBounds.Right;
			topMargin = (int)e.MarginBounds.Top;
			bottomMargin = (int)e.MarginBounds.Bottom;
			InvoiceWidth = (int)e.MarginBounds.Width;
			InvoiceHeight = (int)e.MarginBounds.Height;
			
			if (!ReadInvoice)
				ReadInvoiceData();

			SetInvoiceHead(e.Graphics); // Draw Invoice Head
			SetOrderData(e.Graphics); // Draw Order Data
			SetInvoiceData(e.Graphics, e); // Draw Invoice Data

			ReadInvoice = true;
		}

		private void btnPreview_Click(object sender, System.EventArgs e)
		{
			ReadInvoice = false;
			DisplayInvoice(); // Print Preview
		}

		private void btnDialog_Click(object sender, System.EventArgs e)
		{
			ReadInvoice = false;
			DisplayDialog(); // Print Dialog
		}

		private void btnPrint_Click(object sender, System.EventArgs e)
		{
			ReadInvoice = false;
			PrintReport(); // Print Invoice
		}

		private void btnClose_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}
	}
}
