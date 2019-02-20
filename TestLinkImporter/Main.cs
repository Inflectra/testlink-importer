using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.IO;
using System.Linq;
using System.Collections.Generic;

using Meyn.TestLink;

namespace Inflectra.SpiraTest.AddOns.TestLinkImporter
{
	/// <summary>
	/// This is the code behind class for the utility that imports projects from
	/// TestLink into Inflectra SpiraTest
	/// </summary>
	public class MainForm : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ComboBox cboProject;
		private System.Windows.Forms.TextBox txtLogin;
        private System.Windows.Forms.Button btnCancel;
		private System.Windows.Forms.Button btnAuthenticate;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.TextBox txtServer;
		private System.Windows.Forms.Button btnNext;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		protected ImportForm importForm;
		private System.Windows.Forms.PictureBox pictureBox1;
		private System.Windows.Forms.GroupBox groupBox2;
		public System.Windows.Forms.CheckBox chkImportTestSuites;
		public System.Windows.Forms.CheckBox chkImportTestSteps;
        public System.Windows.Forms.CheckBox chkImportTestCases;
        private CheckBox chkPassword;
		protected ProgressForm progressForm;

        /// <summary>
        /// The id of the selected test link project
        /// </summary>
        public int TestLinkProjectId
        {
            get;
            set;
        }

        /// <summary>
        /// The name of the selected test link project
        /// </summary>
        public string TestLinkProjectName
        {
            get;
            set;
        }

		public MainForm()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			// Add any event handlers
			this.Closing += new CancelEventHandler(MainForm_Closing);

			//Set the initial state of any buttons
			this.btnNext.Enabled = false;

			//Create the other forms and set a handle to this form and the import form
			this.importForm = new ImportForm();
			this.progressForm = new ProgressForm();
			this.importForm.MainFormHandle = this;
			this.importForm.ProgressFormHandle = this.progressForm;
			this.progressForm.MainFormHandle = this;
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.btnNext = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.chkPassword = new System.Windows.Forms.CheckBox();
            this.label6 = new System.Windows.Forms.Label();
            this.txtServer = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.btnAuthenticate = new System.Windows.Forms.Button();
            this.cboProject = new System.Windows.Forms.ComboBox();
            this.txtLogin = new System.Windows.Forms.TextBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.chkImportTestCases = new System.Windows.Forms.CheckBox();
            this.chkImportTestSteps = new System.Windows.Forms.CheckBox();
            this.chkImportTestSuites = new System.Windows.Forms.CheckBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // btnNext
            // 
            this.btnNext.Location = new System.Drawing.Point(451, 74);
            this.btnNext.Name = "btnNext";
            this.btnNext.Size = new System.Drawing.Size(115, 26);
            this.btnNext.TabIndex = 0;
            this.btnNext.Text = "Next >";
            this.btnNext.Click += new System.EventHandler(this.btnNext_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(326, 74);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(106, 26);
            this.btnCancel.TabIndex = 1;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(19, 18);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(528, 27);
            this.label1.TabIndex = 6;
            this.label1.Text = "1. Connect to TestLink";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.chkPassword);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.txtServer);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.btnAuthenticate);
            this.groupBox1.Controls.Add(this.cboProject);
            this.groupBox1.Controls.Add(this.txtLogin);
            this.groupBox1.ForeColor = System.Drawing.Color.Black;
            this.groupBox1.Location = new System.Drawing.Point(29, 55);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(576, 203);
            this.groupBox1.TabIndex = 10;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "TestLink Configuration";
            // 
            // chkPassword
            // 
            this.chkPassword.AutoSize = true;
            this.chkPassword.Location = new System.Drawing.Point(115, 93);
            this.chkPassword.Name = "chkPassword";
            this.chkPassword.Size = new System.Drawing.Size(127, 21);
            this.chkPassword.TabIndex = 22;
            this.chkPassword.Text = "Remember Key";
            this.chkPassword.UseVisualStyleBackColor = true;
            // 
            // label6
            // 
            this.label6.Location = new System.Drawing.Point(29, 28);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(77, 18);
            this.label6.TabIndex = 21;
            this.label6.Text = "Url:";
            this.label6.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // txtServer
            // 
            this.txtServer.Location = new System.Drawing.Point(115, 28);
            this.txtServer.Name = "txtServer";
            this.txtServer.Size = new System.Drawing.Size(403, 22);
            this.txtServer.TabIndex = 20;
            this.txtServer.Text = "http://myserver/testlink";
            // 
            // label5
            // 
            this.label5.Location = new System.Drawing.Point(19, 144);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(154, 19);
            this.label5.TabIndex = 19;
            this.label5.Text = "Project:";
            this.label5.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(10, 65);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(96, 18);
            this.label2.TabIndex = 16;
            this.label2.Text = "API Key:";
            this.label2.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // btnAuthenticate
            // 
            this.btnAuthenticate.Location = new System.Drawing.Point(413, 106);
            this.btnAuthenticate.Name = "btnAuthenticate";
            this.btnAuthenticate.Size = new System.Drawing.Size(105, 27);
            this.btnAuthenticate.TabIndex = 14;
            this.btnAuthenticate.Text = "Authenticate";
            this.btnAuthenticate.Click += new System.EventHandler(this.btnAuthenticate_Click);
            // 
            // cboProject
            // 
            this.cboProject.Location = new System.Drawing.Point(182, 144);
            this.cboProject.Name = "cboProject";
            this.cboProject.Size = new System.Drawing.Size(336, 24);
            this.cboProject.TabIndex = 13;
            // 
            // txtLogin
            // 
            this.txtLogin.Location = new System.Drawing.Point(115, 65);
            this.txtLogin.Name = "txtLogin";
            this.txtLogin.Size = new System.Drawing.Size(403, 22);
            this.txtLogin.TabIndex = 10;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.btnCancel);
            this.groupBox2.Controls.Add(this.chkImportTestCases);
            this.groupBox2.Controls.Add(this.chkImportTestSteps);
            this.groupBox2.Controls.Add(this.chkImportTestSuites);
            this.groupBox2.Controls.Add(this.btnNext);
            this.groupBox2.ForeColor = System.Drawing.Color.Black;
            this.groupBox2.Location = new System.Drawing.Point(29, 274);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(576, 111);
            this.groupBox2.TabIndex = 12;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Import Options";
            // 
            // chkImportTestCases
            // 
            this.chkImportTestCases.Checked = true;
            this.chkImportTestCases.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkImportTestCases.Enabled = false;
            this.chkImportTestCases.Location = new System.Drawing.Point(202, 28);
            this.chkImportTestCases.Name = "chkImportTestCases";
            this.chkImportTestCases.Size = new System.Drawing.Size(220, 27);
            this.chkImportTestCases.TabIndex = 4;
            this.chkImportTestCases.Text = "Test Cases";
            // 
            // chkImportTestSteps
            // 
            this.chkImportTestSteps.Checked = true;
            this.chkImportTestSteps.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkImportTestSteps.Enabled = false;
            this.chkImportTestSteps.Location = new System.Drawing.Point(19, 55);
            this.chkImportTestSteps.Name = "chkImportTestSteps";
            this.chkImportTestSteps.Size = new System.Drawing.Size(221, 28);
            this.chkImportTestSteps.TabIndex = 3;
            this.chkImportTestSteps.Text = "Test Steps";
            this.chkImportTestSteps.CheckedChanged += new System.EventHandler(this.chkImportTestCases_CheckedChanged);
            // 
            // chkImportTestSuites
            // 
            this.chkImportTestSuites.Checked = true;
            this.chkImportTestSuites.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkImportTestSuites.Enabled = false;
            this.chkImportTestSuites.Location = new System.Drawing.Point(19, 28);
            this.chkImportTestSuites.Name = "chkImportTestSuites";
            this.chkImportTestSuites.Size = new System.Drawing.Size(221, 27);
            this.chkImportTestSuites.TabIndex = 2;
            this.chkImportTestSuites.Text = "Test Suites";
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::Inflectra.SpiraTest.AddOns.TestLinkImporter.Properties.Resources.TestLink_Icon;
            this.pictureBox1.Location = new System.Drawing.Point(566, 9);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(48, 46);
            this.pictureBox1.TabIndex = 11;
            this.pictureBox1.TabStop = false;
            // 
            // MainForm
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(6, 15);
            this.ClientSize = new System.Drawing.Size(619, 437);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "MainForm";
            this.Text = "SpiraTest Importer for TestLink";
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);

		}
		#endregion

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main() 
		{
            AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(CurrentDomain_UnhandledException);

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
			Application.Run(new MainForm());
		}

        private static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            MiniDump.CreateMiniDump();
        }

		/// <summary>
		/// Authenticates the user from the providing server/login/password information
		/// </summary>
		/// <param name="sender">The sending object</param>
		/// <param name="e">The event arguments</param>
		private void btnAuthenticate_Click(object sender, System.EventArgs e)
		{
			//Disable the next button
			this.btnNext.Enabled = false;

			//Make sure that a login was entered
			if (this.txtLogin.Text.Trim() == "")
			{
				MessageBox.Show ("You need to enter a login", "Authentication Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				return;
			}

			//Make sure that a server was entered
			if (this.txtServer.Text.Trim() == "")
			{
				MessageBox.Show ("You need to enter the URL to your TestLink instance", "Authentication Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				return;
			}

            try
            {
                //Instantiate the connection to TestLink (by getting a list of projects)
                string apiKey = this.txtLogin.Text.Trim();
                string url = this.txtServer.Text.Trim() + Utils.TESTLINK_URL_SUFFIX_XML_RPC;
                TestLink testLink = new TestLink(apiKey, url);
                List<TestProject> projects = testLink.GetProjects();

                if (projects.Count > 0)
                {
                    MessageBox.Show("You have logged into TestLink Successfully", "Authentication", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("You have logged into TestLink successfully, but you don't have any projects to import!", "No Projects Found", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                //Now we need to populate the list of projects
                List<NameIdObject> projectsList = new List<NameIdObject>();
                //Convert into a standard bindable list source
                for (int i = 0; i < projects.Count; i++)
                {
                    NameIdObject project = new NameIdObject();
                    project.Id = projects[i].id.ToString();
                    project.Name = projects[i].name;
                    projectsList.Add(project);
                }

                //Sort by name
                projectsList = projectsList.OrderBy(p => p.Name).ToList();

                this.cboProject.DisplayMember = "Name";
                this.cboProject.DataSource = projectsList;

                //Enable the Next button
                this.btnNext.Enabled = true;
            }
            catch (Exception exception)
            {
                MessageBox.Show("General error accessing the TestLink API. The error message is: " + exception.Message, "Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }

		/// <summary>
		/// Closes the application
		/// </summary>
		/// <param name="sender">The sending object</param>
		/// <param name="e">The event arguments</param>
		private void btnCancel_Click(object sender, System.EventArgs e)
		{
			//Close the application
			this.Close();
		}

		/// <summary>
		/// Called if the form is closed
		/// </summary>
		/// <param name="sender">The sending object</param>
		/// <param name="e">The event arguments</param>
		private void MainForm_Closing(object sender, CancelEventArgs e)
		{
            //Nothing needs to be done
		}

		/// <summary>
		/// Called when the Next button is clicked. Switches to the second form
		/// </summary>
		/// <param name="sender">The sending object</param>
		/// <param name="e">The event arguments</param>
		private void btnNext_Click(object sender, System.EventArgs e)
		{
            //Store the info in settings for later
            Properties.Settings.Default.TestLinkUrl = this.txtServer.Text.Trim();
            if (chkPassword.Checked)
            {
                Properties.Settings.Default.TestLinkApiKey = this.txtLogin.Text.Trim();
            }
            else
            {
                Properties.Settings.Default.TestLinkApiKey = "";
            }
            Properties.Settings.Default.TestSuites = this.chkImportTestSuites.Checked;
            Properties.Settings.Default.TestSteps = this.chkImportTestSteps.Checked;
            Properties.Settings.Default.TestCases = this.chkImportTestCases.Checked;
            //Properties.Settings.Default.Attachments = this.chkImportAttachments.Checked;
            Properties.Settings.Default.Save();

            //Always put the current API Key in settings after save, for use in current run
            Properties.Settings.Default.TestLinkApiKey = this.txtLogin.Text.Trim();

            //Store the current test rail project id for use by the import thread
            this.TestLinkProjectId = Int32.Parse(((NameIdObject)cboProject.SelectedItem).Id);
            this.TestLinkProjectName = ((NameIdObject)cboProject.SelectedItem).Name;

			//Hide the current form
			this.Hide();

			//Show the second page in the import wizard
			this.importForm.Show();
		}

		/// <summary>
		/// Change the active status of the test run import checkbox depending on this selection
		/// </summary>
		/// <param name="sender">The sending object</param>
		/// <param name="e">The event arguments</param>
		private void chkImportTestCases_CheckedChanged(object sender, System.EventArgs e)
		{
			this.chkImportTestCases.Enabled = this.chkImportTestSteps.Checked;
			this.chkImportTestCases.Checked = false;
        }

        /// <summary>
        /// Populates the fields when the form is loaded
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MainForm_Load(object sender, EventArgs e)
        {
            this.txtServer.Text = Properties.Settings.Default.TestLinkUrl;
            if (String.IsNullOrEmpty(Properties.Settings.Default.TestLinkApiKey))
            {
                this.chkPassword.Checked = false;
                this.txtLogin.Text = "";
            }
            else
            {
                this.chkPassword.Checked = true;
                this.txtLogin.Text = Properties.Settings.Default.TestLinkApiKey;
            }

            this.chkImportTestSuites.Checked = Properties.Settings.Default.TestSuites;
            this.chkImportTestSteps.Checked = Properties.Settings.Default.TestSteps;
            this.chkImportTestCases.Checked = Properties.Settings.Default.TestCases;
        }
	}
}
