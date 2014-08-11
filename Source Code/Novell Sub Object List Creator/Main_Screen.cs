using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Novell.Directory.Ldap;
using vbAccelerator.Components.Shell;



namespace Novell_Sub_Object_List_Creator
{
	/// <summary>
	/// Summary description for Main_Screen.
	/// </summary>
	public class Main_Screen : System.Windows.Forms.Form
	{
		private SysTray sysTray;

		private System.Windows.Forms.RichTextBox label1;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Panel panel2;
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.Label Statusbar;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.TextBox SearchContext;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.TextBox SearchParameter;
		private System.Windows.Forms.Label resultsreturned;
		private System.Windows.Forms.ImageList ilsIcons;
		private System.Windows.Forms.ComboBox comboBox1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Panel panel1;
		private System.Windows.Forms.Button button2;
		private System.Windows.Forms.TextBox savefilename;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		private System.Windows.Forms.PictureBox pictureBox1;
		private System.Windows.Forms.Label label3;
		private System.ComponentModel.IContainer components;

		public Main_Screen()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
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
			this.components = new System.ComponentModel.Container();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(Main_Screen));
			this.label1 = new System.Windows.Forms.RichTextBox();
			this.label5 = new System.Windows.Forms.Label();
			this.panel2 = new System.Windows.Forms.Panel();
			this.button1 = new System.Windows.Forms.Button();
			this.Statusbar = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.SearchContext = new System.Windows.Forms.TextBox();
			this.label9 = new System.Windows.Forms.Label();
			this.label10 = new System.Windows.Forms.Label();
			this.SearchParameter = new System.Windows.Forms.TextBox();
			this.resultsreturned = new System.Windows.Forms.Label();
			this.ilsIcons = new System.Windows.Forms.ImageList(this.components);
			this.comboBox1 = new System.Windows.Forms.ComboBox();
			this.label2 = new System.Windows.Forms.Label();
			this.panel1 = new System.Windows.Forms.Panel();
			this.button2 = new System.Windows.Forms.Button();
			this.savefilename = new System.Windows.Forms.TextBox();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.pictureBox1 = new System.Windows.Forms.PictureBox();
			this.label3 = new System.Windows.Forms.Label();
			this.SuspendLayout();
			// 
			// label1
			// 
			this.label1.BackColor = System.Drawing.Color.WhiteSmoke;
			this.label1.BorderStyle = System.Windows.Forms.BorderStyle.None;
			this.label1.ForeColor = System.Drawing.Color.Black;
			this.label1.Location = new System.Drawing.Point(24, 233);
			this.label1.Name = "label1";
			this.label1.ReadOnly = true;
			this.label1.Size = new System.Drawing.Size(511, 94);
			this.label1.TabIndex = 5;
			this.label1.Text = "";
			// 
			// label5
			// 
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 6F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label5.ForeColor = System.Drawing.Color.Black;
			this.label5.Location = new System.Drawing.Point(24, 12);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(472, 36);
			this.label5.TabIndex = 7;
			this.label5.Text = "TO RUN A SEARCH ON THE NOVELL OBJECTS CONTAINED ON THE UCT TREE, SIMPLY ENTER YOU" +
				"R SEARCH PARAMETERS AS INDICATED AND HIT THE PROCESS BUTTON TO RETRIEVE THE RESU" +
				"LTS. YOU ALSO NEED TO SELECT A VALID FILENAME TO SAVE THE SEARCH RESULTS TO.";
			// 
			// panel2
			// 
			this.panel2.BackColor = System.Drawing.Color.WhiteSmoke;
			this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.panel2.ForeColor = System.Drawing.Color.Black;
			this.panel2.Location = new System.Drawing.Point(23, 232);
			this.panel2.Name = "panel2";
			this.panel2.Size = new System.Drawing.Size(513, 96);
			this.panel2.TabIndex = 9;
			// 
			// button1
			// 
			this.button1.BackColor = System.Drawing.Color.LightGray;
			this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
			this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button1.ForeColor = System.Drawing.Color.Black;
			this.button1.Location = new System.Drawing.Point(432, 152);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(104, 40);
			this.button1.TabIndex = 4;
			this.button1.Text = "PROCESS";
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// Statusbar
			// 
			this.Statusbar.Font = new System.Drawing.Font("Microsoft Sans Serif", 6F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.Statusbar.ForeColor = System.Drawing.Color.Firebrick;
			this.Statusbar.Location = new System.Drawing.Point(24, 382);
			this.Statusbar.Name = "Statusbar";
			this.Statusbar.Size = new System.Drawing.Size(400, 9);
			this.Statusbar.TabIndex = 11;
			this.Statusbar.TextAlign = System.Drawing.ContentAlignment.BottomLeft;
			// 
			// label6
			// 
			this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 6F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label6.ForeColor = System.Drawing.Color.Gray;
			this.label6.Location = new System.Drawing.Point(24, 372);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(400, 9);
			this.label6.TabIndex = 12;
			this.label6.TextAlign = System.Drawing.ContentAlignment.BottomLeft;
			// 
			// label7
			// 
			this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 6F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label7.ForeColor = System.Drawing.Color.Gray;
			this.label7.Location = new System.Drawing.Point(24, 362);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(400, 9);
			this.label7.TabIndex = 13;
			this.label7.TextAlign = System.Drawing.ContentAlignment.BottomLeft;
			// 
			// label8
			// 
			this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 6F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label8.ForeColor = System.Drawing.Color.Gray;
			this.label8.Location = new System.Drawing.Point(24, 352);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(400, 9);
			this.label8.TabIndex = 14;
			this.label8.TextAlign = System.Drawing.ContentAlignment.BottomLeft;
			// 
			// SearchContext
			// 
			this.SearchContext.BackColor = System.Drawing.Color.WhiteSmoke;
			this.SearchContext.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.SearchContext.ForeColor = System.Drawing.Color.Firebrick;
			this.SearchContext.Location = new System.Drawing.Point(24, 72);
			this.SearchContext.Name = "SearchContext";
			this.SearchContext.Size = new System.Drawing.Size(272, 20);
			this.SearchContext.TabIndex = 15;
			this.SearchContext.Text = "Students.com.main.uct";
			// 
			// label9
			// 
			this.label9.Font = new System.Drawing.Font("Microsoft Sans Serif", 6F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label9.ForeColor = System.Drawing.Color.Firebrick;
			this.label9.Location = new System.Drawing.Point(24, 56);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(112, 16);
			this.label9.TabIndex = 16;
			this.label9.Text = "SEARCH CONTEXT";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label10
			// 
			this.label10.Font = new System.Drawing.Font("Microsoft Sans Serif", 6F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label10.ForeColor = System.Drawing.Color.Firebrick;
			this.label10.Location = new System.Drawing.Point(24, 104);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(112, 16);
			this.label10.TabIndex = 18;
			this.label10.Text = "SEARCH PARAMETER";
			this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// SearchParameter
			// 
			this.SearchParameter.BackColor = System.Drawing.Color.WhiteSmoke;
			this.SearchParameter.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.SearchParameter.ForeColor = System.Drawing.Color.Firebrick;
			this.SearchParameter.Location = new System.Drawing.Point(24, 120);
			this.SearchParameter.Name = "SearchParameter";
			this.SearchParameter.Size = new System.Drawing.Size(272, 20);
			this.SearchParameter.TabIndex = 17;
			this.SearchParameter.Text = "CN=*";
			// 
			// resultsreturned
			// 
			this.resultsreturned.Font = new System.Drawing.Font("Microsoft Sans Serif", 6F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.resultsreturned.ForeColor = System.Drawing.Color.Firebrick;
			this.resultsreturned.Location = new System.Drawing.Point(24, 208);
			this.resultsreturned.Name = "resultsreturned";
			this.resultsreturned.Size = new System.Drawing.Size(232, 16);
			this.resultsreturned.TabIndex = 19;
			this.resultsreturned.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ilsIcons
			// 
			this.ilsIcons.ImageSize = new System.Drawing.Size(16, 16);
			this.ilsIcons.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("ilsIcons.ImageStream")));
			this.ilsIcons.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// comboBox1
			// 
			this.comboBox1.BackColor = System.Drawing.Color.WhiteSmoke;
			this.comboBox1.ForeColor = System.Drawing.Color.DimGray;
			this.comboBox1.Items.AddRange(new object[] {
														   "Students.com.main.uct",
														   "Staff.com.main.uct",
														   "com.main.uct"});
			this.comboBox1.Location = new System.Drawing.Point(384, 88);
			this.comboBox1.Name = "comboBox1";
			this.comboBox1.Size = new System.Drawing.Size(144, 21);
			this.comboBox1.TabIndex = 20;
			this.comboBox1.Text = "Students.com.main.uct";
			this.comboBox1.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
			// 
			// label2
			// 
			this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 6F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label2.ForeColor = System.Drawing.Color.Black;
			this.label2.Location = new System.Drawing.Point(384, 72);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(144, 16);
			this.label2.TabIndex = 21;
			this.label2.Text = "CONTEXT QUICK SELECTOR";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// panel1
			// 
			this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.panel1.Location = new System.Drawing.Point(376, 62);
			this.panel1.Name = "panel1";
			this.panel1.Size = new System.Drawing.Size(160, 56);
			this.panel1.TabIndex = 22;
			// 
			// button2
			// 
			this.button2.BackColor = System.Drawing.Color.RosyBrown;
			this.button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
			this.button2.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.button2.Location = new System.Drawing.Point(304, 168);
			this.button2.Name = "button2";
			this.button2.Size = new System.Drawing.Size(56, 20);
			this.button2.TabIndex = 44;
			this.button2.Text = "BROWSE";
			this.button2.Click += new System.EventHandler(this.button2_Click);
			// 
			// savefilename
			// 
			this.savefilename.BackColor = System.Drawing.Color.WhiteSmoke;
			this.savefilename.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.savefilename.ForeColor = System.Drawing.Color.Black;
			this.savefilename.Location = new System.Drawing.Point(24, 168);
			this.savefilename.Name = "savefilename";
			this.savefilename.ReadOnly = true;
			this.savefilename.Size = new System.Drawing.Size(272, 20);
			this.savefilename.TabIndex = 42;
			this.savefilename.Text = "";
			// 
			// saveFileDialog1
			// 
			this.saveFileDialog1.DefaultExt = "txt";
			this.saveFileDialog1.FileName = "ActivityLog.txt";
			this.saveFileDialog1.Filter = "Text files|*.txt";
			this.saveFileDialog1.Title = "Search Results File ";
			// 
			// pictureBox1
			// 
			this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
			this.pictureBox1.Location = new System.Drawing.Point(496, 0);
			this.pictureBox1.Name = "pictureBox1";
			this.pictureBox1.Size = new System.Drawing.Size(96, 176);
			this.pictureBox1.TabIndex = 45;
			this.pictureBox1.TabStop = false;
			// 
			// label3
			// 
			this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 6F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label3.ForeColor = System.Drawing.Color.Firebrick;
			this.label3.Location = new System.Drawing.Point(24, 152);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(112, 16);
			this.label3.TabIndex = 46;
			this.label3.Text = "OBJECT RESULTS FILE";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// Main_Screen
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.BackColor = System.Drawing.Color.White;
			this.ClientSize = new System.Drawing.Size(560, 396);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.button2);
			this.Controls.Add(this.savefilename);
			this.Controls.Add(this.SearchParameter);
			this.Controls.Add(this.SearchContext);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.comboBox1);
			this.Controls.Add(this.resultsreturned);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.Statusbar);
			this.Controls.Add(this.button1);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.panel2);
			this.Controls.Add(this.panel1);
			this.Controls.Add(this.pictureBox1);
			this.ForeColor = System.Drawing.Color.Black;
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.MaximizeBox = false;
			this.MaximumSize = new System.Drawing.Size(568, 430);
			this.MinimumSize = new System.Drawing.Size(568, 430);
			this.Name = "Main_Screen";
			this.Text = "Novell Sub Object List Creator 20060309.1";
			this.Load += new System.EventHandler(this.Main_Screen_Load);
			this.ResumeLayout(false);

		}
		#endregion

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main() 
		{
			Application.Run(new Main_Screen());
		}

		protected override void OnHandleCreated(EventArgs e)
		{
			if (null != sysTray)
			{
				sysTray.AssignHandle(this.Handle);	
				sysTray.ShowInSysTray = true;
			}
			base.OnHandleCreated(e);
		}

		private void sysTray_DoubleClick(object sender, EventArgs e)
		{
			RestoreMe();
		}

		/// <summary>
		/// Performs the following steps as necessary:
		/// unhides, unmiminizes, brings to foreground and
		/// sets as active application.
		/// </summary>
		private void RestoreMe()
		{
			if (!this.Visible)
			{
				this.Visible = true;
			}
			if (this.WindowState == FormWindowState.Minimized)
			{
				this.WindowState = FormWindowState.Normal;
			}
			this.BringToFront();
		}

		private void Main_Screen_Load(object sender, System.EventArgs e)
		{
			try
			{
//				Context.SelectedIndex = 1;
//				if (Context.SelectedIndex > -1)
//				{
//					SearchContext.Text = Context.Items[Context.SelectedIndex].ToString();
//					SearchContext.Refresh();
//				}




				// Assign the SysTray object to this form
				sysTray = new SysTray(this);
				// Set the Icon:
				sysTray.IconImageList = ilsIcons;
				sysTray.IconIndex = 0;
				// Set the default context menu:
				//sysTray.Menu = mnuContext;
				// Set the tooltip text:
				sysTray.ToolTipText = "Novell Sub Object List Creator";
			
				// events:
				/*sysTray.MouseDown += new MouseEventHandler(sysTray_MouseDown);
				sysTray.MouseUp += new MouseEventHandler(sysTray_MouseUp);
				sysTray.MouseMove += new MouseEventHandler(sysTray_MouseMove);*/
				sysTray.Click += new EventHandler(sysTray_DoubleClick);
				sysTray.DoubleClick += new EventHandler(sysTray_DoubleClick);
				/*sysTray.KeySelect += new EventHandler(sysTray_KeySelect);
				sysTray.EnterSelect += new EventHandler(sysTray_EnterSelect);
				sysTray.BalloonClicked += new EventHandler(sysTray_BalloonClicked);
				sysTray.BalloonHide += new EventHandler(sysTray_BalloonHide);
				sysTray.BalloonShow += new EventHandler(sysTray_BalloonShow);
				sysTray.BalloonTimeOut += new EventHandler(sysTray_BalloonTimeOut);
*/
				// Show:
				sysTray.ShowInSysTray = true;


			}
			catch(System.Exception except)
			{
				Error_Handler(except);
			}
		}

		private void Error_Handler(System.Exception except)
		{
			label1_Message("--------------\nError Encountered:\n" + except.ToString());
			Statusbar_Message("Error Encountered");
		}

		private void Error_Handler(string except)
		{
			label1_Message("--------------\nError Encountered:\n" + except);
			Statusbar_Message("Error Encountered");
		}

		private void Statusbar_Message(string message)
		{
			label8.Text = label7.Text;
			label8.Refresh();
			label7.Text = label6.Text;
			label7.Refresh();
			label6.Text = Statusbar.Text;
			label6.Refresh();
			message = message.ToUpper();
			Statusbar.Text = message;
			Statusbar.Refresh();
		}
	
		private void label1_Message(string message)
		{
			if (label1.TextLength > 0)
				label1.Text = label1.Text + "\n" + message;
			else
				label1.Text = label1.Text  + message;
			label1.Refresh();
		}
		
		private int ldap_check(int serveruse)
		{
			try
			{
//				if (serveruse == 1)
//			Statusbar_Message("Retrieving username input");
//				if (Username.Text.Equals("") == true)
//				{
//					Error_Handler("Input Error: Username field has been left blank.");
//					Statusbar_Message("Operation failed");
//					return 0;
//				}
//if (serveruse == 1)
//				Statusbar_Message("Retrieving password input");
//				string userPasswd = Password.Text;
//				
//				if (Password.Text.Equals("") == true)
//				{
//					Error_Handler("Input Error: Password field has been left blank.");
//					Statusbar_Message("Operation failed");
//					return 0;
//				}	
//
//				if (Context.SelectedIndex < 0)
//				{
//					Error_Handler("Invalid Context selected from the list.");
//					Statusbar_Message("Operation failed");
//					return 0;
//				}	
//
//				string tdn = "";
//				string[] tempdn = Context.Items[Context.SelectedIndex].ToString().Split('.');
//				System.Collections.IEnumerator runner = tempdn.GetEnumerator();
//				while(runner.MoveNext())
//				{
//					if (runner.Current.Equals("uct") == false)
//						tdn = tdn + "OU="+ runner.Current + ",";
//					else
//						tdn = tdn + "O="+ runner.Current + ",";
//				}
//				tdn = tdn.Remove(tdn.Length-1,1);
//				string userDN = "CN=" + Username.Text.ToUpper() + "," + tdn;
//				if (serveruse == 1)
//				Statusbar_Message("Setting DN string as " + userDN);
//
//			
				
				string ldapHost;
				int ldapPort;
//				int ldapVersion = LdapConnection.Ldap_V3;
				if (serveruse == 1)
				{
					ldapHost = "rep1.uct.ac.za";
					ldapPort = LdapConnection.DEFAULT_PORT;
				}
				else
				{
					ldapHost = "rep2.uct.ac.za";
					ldapPort = LdapConnection.DEFAULT_PORT;
				}
				Statusbar_Message("Setting LDAP server as " + ldapHost + ":" + ldapPort);

			//	Statusbar_Message("Attempting to bind " + Username.Text + " to " + ldapHost + ":" + ldapPort);
				// Creating an LdapConnection instance
				LdapConnection ldapConn= new LdapConnection();
				try				
				{
					//Connect function will create a socket
					//ldapConn.SecureSocketLayer = true;
					ldapConn.Connect(ldapHost,ldapPort);
					//Bind function will Bind the user
					//ldapConn.startTLS();
				//	ldapConn.Bind(ldapVersion,userDN,userPasswd);
					//ldapConn.stopTLS();
								}
				catch(System.Exception except)
				{
					Error_Handler(except);
					if (serveruse == 1)
					{
						Statusbar_Message("Contacting Secondary LDAP server");
						ldap_check(2);
						return 0;
					}
					else
					{
						Statusbar_Message("Operation failed");
						return 0;
						
					}
				}

				
				//LdapAttribute attr = new LdapAttribute("userPassword", userPasswd);
				//bool correct = ldapConn.Compare(userDN, attr);
				//label1_Message("Password Verify: " + correct);

				string stdn = "";
				string[] stempdn = SearchContext.Text.Split('.');
				System.Collections.IEnumerator srunner = stempdn.GetEnumerator();
				while(srunner.MoveNext())
				{
					if (srunner.Current.Equals("uct") == false)
						stdn = stdn + "OU="+ srunner.Current + ",";
					else
						stdn = stdn + "O="+ srunner.Current + ",";
				}
				stdn = stdn.Remove(stdn.Length-1,1);
				Statusbar_Message("Search DN string set as " + stdn);
				if (stdn.Equals("") == true)
				{
					Error_Handler("Input Error: Search DN field has been left blank.");
					Statusbar_Message("Operation failed");
					return 0;
				}

				Statusbar_Message("Search Parameter string set as " + SearchParameter.Text);
				if (SearchParameter.Text.Equals("") == true)
				{
					Error_Handler("Input Error: Search Parameter field has been left blank.");
					Statusbar_Message("Operation failed");
					return 0;
				}

				Statusbar_Message("Retrieving Object Details");
				resultsreturned.Text = "0";
				resultsreturned.Text = resultsreturned.Text + " RESULTS RETURNED";
				resultsreturned.Refresh();
				bool LoopsOn;
				LoopsOn = false;
				if (SearchParameter.Text.ToUpper() == "CN=*")
				{
					LoopsOn = true;
				}
				else
				{
					LoopsOn = false;
				}
				

				string[] searchparams;
				if (LoopsOn == false)
				{
					searchparams = new string[1]{SearchParameter.Text};
				}
				else
				{
					searchparams = new string[702];
					string[] outerloop = {"A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"};
					string[] innerloop = {"-","A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"};
					//searchparams = new string[676];
					//string[] outerloop = {"A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"};
					//string[] innerloop = {"A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"};

					int current = 0;
					foreach(string outer in outerloop)
					{
						string c = outer;
						foreach(string inner in innerloop)
						{
							c = "CN=" + outer +  inner + "*";
							searchparams[current] = c;
							//label1_Message(current + ": " + c);
							current = current + 1;
						}
						
					}
				}
				//MessageBox.Show(searchparams.Length.ToString());

				foreach(string arrsearchterm in searchparams)
				{

					// Searches in the Marketing container and return all child entries just below this
					//container i.e Single level search
					LdapSearchQueue queue=ldapConn.Search(stdn,LdapConnection.SCOPE_SUB,
						//"objectClass=*",
						//"CN=" + Username.Text.ToUpper(),				
						arrsearchterm,
						null,		
						false,
						(LdapSearchQueue) null,
						(LdapSearchConstraints) null );

					LdapMessage message;
				
					while ((message = queue.getResponse()) != null)
				
					{
						if (message is LdapSearchResult)
						{
							double doh = (Double.Parse(resultsreturned.Text.Replace(" RESULTS RETURNED","")) + 1);
							resultsreturned.Text = doh.ToString();
							resultsreturned.Text = resultsreturned.Text + " RESULTS RETURNED";
							resultsreturned.Refresh();
							LdapEntry entry = ((LdapSearchResult) message).Entry;
							//label1_Message("------------------------------------------------\n ENTRY:\n------------------------------------------------");
							////	label1_Message(" " + entry.DN );
							string writemessage = "";
								writemessage = entry.DN.ToString().ToLower().Replace("cn=",".").Replace("o=",".").Replace("ou=",".").Replace(",","");
							ActivityLogger(writemessage);

							


						
						}
					}
				}
Statusbar_Message("Disconnecting from LDAP server");

				ldapConn.Disconnect();
				Statusbar_Message("Operation successfully completed");
				sysTray.ShowBalloonTip("Operation successfully completed. " + resultsreturned.Text.ToLower() + ".", NotifyIconBalloonIconFlags.NIIF_INFO, "Novell Sub Object List Creator");
			
				return 1;
			}
			catch(System.Exception except)
			{
								Error_Handler(except);
				Statusbar_Message("Operation failed");
			}
			return 0;
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
			if (savefilename.Text.Length > 0)
			{
				label8.Text = "";
				label8.Refresh();
				label7.Text = "";
				label7.Refresh();
				label6.Text = "";
				label6.Refresh();
				Statusbar.Text = "";
				Statusbar.Refresh();
				label1.Text = "";
				label1.Refresh();
				ldap_check(1);
			}
			else
			{
				MessageBox.Show("Please ensure you have a valid save file name entered into the Results File textbox.");
			}
		}

		private void ActivityLogger(string Message)
		{
			try
			{
				if (savefilename.Text.Length > 0)
				{

					System.IO.StreamWriter filewriter;
					filewriter = new System.IO.StreamWriter(savefilename.Text,true);
					filewriter.WriteLine(Message);
					filewriter.Flush();
					filewriter.Close();
				}
			}
			catch(System.Exception except)
			{
				Error_Handler("ActivityLogger: " + except);
				Statusbar_Message("Activity Logging failed");
			}
		}

		private void comboBox1_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			SearchContext.Text = comboBox1.SelectedItem.ToString();
		}

		private void button2_Click(object sender, System.EventArgs e)
		{
			saveFileDialog1.FileName = "Search_Results_" + String.Format("{0:yyyyMMdd}", DateTime.Now) + ".txt";
			DialogResult result = saveFileDialog1.ShowDialog();
			if (result == DialogResult.OK)
			{
				savefilename.Text = saveFileDialog1.FileName;
			}
		}

//		private void Context_SelectedIndexChanged(object sender, System.EventArgs e)
//		{
//			try
//			{
//				if (Context.SelectedIndex > -1)
//				{
//					SearchContext.Text = Context.Items[Context.SelectedIndex].ToString();
//					SearchContext.Refresh();
//				}
//			}
//			catch(System.Exception except)
//			{
//				Error_Handler(except);
//				Statusbar_Message("Operation failed");
//			}
//			
//		}


	}
}
