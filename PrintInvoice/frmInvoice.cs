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
	/// Summary description for Form1.
	/// </summary>
	public class frmInvoice : System.Windows.Forms.Form
	{
		private System.Windows.Forms.DataGrid datGrid;
		private System.Windows.Forms.Button btnLoad;
		private System.Windows.Forms.Button btnFind;
		private System.Windows.Forms.Button btnExit;
		private System.ComponentModel.Container components = null;

		private string strCon;

		public frmInvoice()
		{
			InitializeComponent();
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if (components != null) 
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
			this.btnExit = new System.Windows.Forms.Button();
			this.btnLoad = new System.Windows.Forms.Button();
			this.datGrid = new System.Windows.Forms.DataGrid();
			this.btnFind = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.datGrid)).BeginInit();
			this.SuspendLayout();
			// 
			// btnExit
			// 
			this.btnExit.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnExit.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.btnExit.Location = new System.Drawing.Point(448, 256);
			this.btnExit.Name = "btnExit";
			this.btnExit.Size = new System.Drawing.Size(88, 20);
			this.btnExit.TabIndex = 33;
			this.btnExit.Text = "Exit";
			this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
			// 
			// btnLoad
			// 
			this.btnLoad.Location = new System.Drawing.Point(16, 256);
			this.btnLoad.Name = "btnLoad";
			this.btnLoad.Size = new System.Drawing.Size(88, 20);
			this.btnLoad.TabIndex = 34;
			this.btnLoad.Text = "Load Data";
			this.btnLoad.Click += new System.EventHandler(this.btnLoad_Click);
			// 
			// datGrid
			// 
			this.datGrid.DataMember = "";
			this.datGrid.HeaderForeColor = System.Drawing.SystemColors.ControlText;
			this.datGrid.Location = new System.Drawing.Point(16, 8);
			this.datGrid.Name = "datGrid";
			this.datGrid.Size = new System.Drawing.Size(520, 240);
			this.datGrid.TabIndex = 35;
			// 
			// btnFind
			// 
			this.btnFind.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnFind.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.btnFind.Location = new System.Drawing.Point(112, 256);
			this.btnFind.Name = "btnFind";
			this.btnFind.Size = new System.Drawing.Size(88, 20);
			this.btnFind.TabIndex = 40;
			this.btnFind.Text = "Find Order";
			this.btnFind.Click += new System.EventHandler(this.btnFind_Click);
			// 
			// frmInvoice
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(550, 283);
			this.Controls.Add(this.btnFind);
			this.Controls.Add(this.datGrid);
			this.Controls.Add(this.btnLoad);
			this.Controls.Add(this.btnExit);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
			this.MaximizeBox = false;
			this.Name = "frmInvoice";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Orders";
			this.Load += new System.EventHandler(this.frmInvoice_Load);
			((System.ComponentModel.ISupportInitialize)(this.datGrid)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		[STAThread]
		static void Main() 
		{
			Application.Run(new frmInvoice());
		}

		private void LoadData()
		{
			// following lines to connect with access database file 'Northwind.mdb' 
			//string MyPass = "";
			//string MyDataFile = Application.StartupPath + @"\DataFile\Northwind.mdb";
            //strCon = @"provider=microsoft.jet.oledb.4.0;data source=" + 
            //    MyDataFile + ";" + "Jet OLEDB:Database Password=" + MyPass + ";";
            strCon = "provider=sqloledb;Data Source=uatdb24.kotakseconline.com;Initial Catalog=KSCS_DB;User ID=kscs_appowner;Password=kscs_appowner@123";

			/* If you are using SQL Server, please replace previous lines with following
			strCon = @"provider=sqloledb;Data Source=PC;Initial Catalog=" +
				"Northwind;Integrated Security=SSPI" + ";";
			and replace 'Data Source=PC' with the name of your system */

			try
			{
				// Get data from tables: Orders, Customers, Employees, Products, Order Details:
				string InvSql = "SELECT  top 20 * from mstestdetails";

				//create an OleDbDataAdapter
				OleDbDataAdapter datAdp = new OleDbDataAdapter(InvSql, strCon);

				//create a command builder
				OleDbCommandBuilder cBuilder = new OleDbCommandBuilder(datAdp);

				//create a DataTable to hold the query results
				DataTable dTable = new DataTable();

				//fill the DataTable
				datAdp.Fill(dTable);

				//set DataSource of DataGrid 
				datGrid.DataSource = dTable;
			}
			catch(Exception e)
			{ 
				MessageBox.Show(e.ToString());
			}
		}

		private void frmInvoice_Load(object sender, System.EventArgs e)
		{
			datGrid.CaptionText = "Orders...";
			btnFind.Enabled = false;
		}

		private void btnLoad_Click(object sender, System.EventArgs e)
		{
			LoadData();
			btnFind.Enabled = true;
			btnLoad.Enabled = false;
		}

		private void btnFind_Click(object sender, System.EventArgs e)
		{
			PrintInvoice.frmInput fInput = new PrintInvoice.frmInput();
			fInput.ShowDialog();
			
			if(frmInput.InvoiceOrder == "")
				return;

			PrintInvoice.frmOrder fOrder = new PrintInvoice.frmOrder();
			fOrder.ShowDialog();
		}

		private void btnExit_Click(object sender, System.EventArgs e)
		{
			Application.Exit();
		}
	}
}
