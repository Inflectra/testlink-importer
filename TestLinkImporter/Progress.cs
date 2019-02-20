using System;
using System.Drawing;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Runtime.InteropServices;
using System.IO;
using System.Net;
using System.Threading;
using System.Text;
using System.Text.RegularExpressions;

using Inflectra.SpiraTest.AddOns.TestLinkImporter.SpiraSoapService;
using System.ServiceModel;
using System.Xml;

namespace Inflectra.SpiraTest.AddOns.TestLinkImporter
{
    /// <summary>
    /// This is the code behind class for the utility that imports projects from
    /// HP Mercury Quality Center / TestDirector into Inflectra SpiraTest
    /// </summary>
    public class ProgressForm : System.Windows.Forms.Form
    {
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button btnCancel;

        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.Container components = null;
        protected MainForm mainForm;
        protected ImportForm importForm;
        private System.Windows.Forms.Button btnExit;
        private System.Windows.Forms.Label lblProgress1;
        private ProgressBar progressBar1;
        private Label lblProgress2;
        private Label lblProgress6;
        private Label lblProgress5;
        private Label lblProgress4;
        private Label lblProgress3;


        #region Properties

        public MainForm MainFormHandle
        {
            get
            {
                return this.mainForm;
            }
            set
            {
                this.mainForm = value;
            }
        }

        #endregion

        public ProgressForm()
        {
            //
            // Required for Windows Form Designer support
            //
            InitializeComponent();

            // Add any event handlers
            this.Closing += new CancelEventHandler(ProgressForm_Closing);
        }

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code
        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ProgressForm));
            this.btnExit = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.lblProgress2 = new System.Windows.Forms.Label();
            this.lblProgress1 = new System.Windows.Forms.Label();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.lblProgress6 = new System.Windows.Forms.Label();
            this.lblProgress5 = new System.Windows.Forms.Label();
            this.lblProgress4 = new System.Windows.Forms.Label();
            this.lblProgress3 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnExit
            // 
            this.btnExit.Location = new System.Drawing.Point(476, 368);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(112, 29);
            this.btnExit.TabIndex = 0;
            this.btnExit.Text = "Done";
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(364, 368);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(103, 29);
            this.btnCancel.TabIndex = 1;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(19, 21);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(494, 29);
            this.label1.TabIndex = 6;
            this.label1.Text = "3. Import in Progress...";
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::Inflectra.SpiraTest.AddOns.TestLinkImporter.Properties.Resources.SpiraTest_Importer_Icon_Large;
            this.pictureBox1.Location = new System.Drawing.Point(551, 10);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(46, 52);
            this.pictureBox1.TabIndex = 7;
            this.pictureBox1.TabStop = false;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.lblProgress2);
            this.groupBox1.Controls.Add(this.lblProgress6);
            this.groupBox1.Controls.Add(this.lblProgress5);
            this.groupBox1.Controls.Add(this.lblProgress4);
            this.groupBox1.Controls.Add(this.lblProgress3);
            this.groupBox1.Controls.Add(this.lblProgress1);
            this.groupBox1.ForeColor = System.Drawing.Color.Black;
            this.groupBox1.Location = new System.Drawing.Point(28, 78);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(560, 237);
            this.groupBox1.TabIndex = 10;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Import Progress";
            // 
            // lblProgress2
            // 
            this.lblProgress2.Location = new System.Drawing.Point(47, 57);
            this.lblProgress2.Name = "lblProgress2";
            this.lblProgress2.Size = new System.Drawing.Size(429, 20);
            this.lblProgress2.TabIndex = 27;
            this.lblProgress2.Text = ">>    Test Suites Imported";
            // 
            // lblProgress1
            // 
            this.lblProgress1.Location = new System.Drawing.Point(47, 31);
            this.lblProgress1.Name = "lblProgress1";
            this.lblProgress1.Size = new System.Drawing.Size(429, 20);
            this.lblProgress1.TabIndex = 26;
            this.lblProgress1.Text = ">>    Project Created";
            // 
            // progressBar1
            // 
            this.progressBar1.ForeColor = System.Drawing.Color.OrangeRed;
            this.progressBar1.Location = new System.Drawing.Point(-3, 345);
            this.progressBar1.Maximum = 6;
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(621, 12);
            this.progressBar1.TabIndex = 28;
            // 
            // lblProgress6
            // 
            this.lblProgress6.Location = new System.Drawing.Point(47, 175);
            this.lblProgress6.Name = "lblProgress6";
            this.lblProgress6.Size = new System.Drawing.Size(429, 20);
            this.lblProgress6.TabIndex = 25;
            this.lblProgress6.Text = ">>    Test Runs Imported";
            // 
            // lblProgress5
            // 
            this.lblProgress5.Location = new System.Drawing.Point(47, 144);
            this.lblProgress5.Name = "lblProgress5";
            this.lblProgress5.Size = new System.Drawing.Size(429, 21);
            this.lblProgress5.TabIndex = 24;
            this.lblProgress5.Text = ">>    Test Plans Imported";
            // 
            // lblProgress4
            // 
            this.lblProgress4.Location = new System.Drawing.Point(47, 113);
            this.lblProgress4.Name = "lblProgress4";
            this.lblProgress4.Size = new System.Drawing.Size(429, 21);
            this.lblProgress4.TabIndex = 22;
            this.lblProgress4.Text = ">>    Import Completed";
            // 
            // lblProgress3
            // 
            this.lblProgress3.Location = new System.Drawing.Point(47, 84);
            this.lblProgress3.Name = "lblProgress3";
            this.lblProgress3.Size = new System.Drawing.Size(429, 20);
            this.lblProgress3.TabIndex = 21;
            this.lblProgress3.Text = ">>    Test Cases Imported";
            // 
            // ProgressForm
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(7, 18);
            this.ClientSize = new System.Drawing.Size(615, 404);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnExit);
            this.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "ProgressForm";
            this.Text = "SpiraTest Importer for TestLink";
            this.Load += new System.EventHandler(this.ProgressForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.ResumeLayout(false);

        }
        #endregion

        /// <summary>
        /// Closes the application
        /// </summary>
        /// <param name="sender">The sending object</param>
        /// <param name="e">The event arguments</param>
        private void btnCancel_Click(object sender, System.EventArgs e)
        {
            //Hide the current form
            this.Hide();

            //Return to the main form
            MainFormHandle.Show();
        }

        /// <summary>
        /// Closes the application
        /// </summary>
        /// <param name="sender">The sending object</param>
        /// <param name="e">The event arguments</param>
        private void btnExit_Click(object sender, System.EventArgs e)
        {
            //Close the application
            this.MainFormHandle.Close();
        }

        /// <summary>
        /// Called if the form is closed
        /// </summary>
        /// <param name="sender">The sending object</param>
        /// <param name="e">The event arguments</param>
        private void ProgressForm_Closing(object sender, CancelEventArgs e)
        {
        }

        /// <summary>
        /// Starts the background thread for importing the data
        /// </summary>
        public void StartImport()
        {
            //Set the initial state of any buttons
            this.btnCancel.Enabled = true;
            this.btnExit.Enabled = false;

            //Clear the progress labels
            ProgressForm_OnProgressUpdate(0);

            //First change the cursor to an hourglass
            this.Cursor = Cursors.WaitCursor;

            //Start the background thread that performs the import
            ImportThread importThread = new ImportThread(this, mainForm.TestLinkProjectId, mainForm.TestLinkProjectName);
            ThreadPool.QueueUserWorkItem(new WaitCallback(importThread.ImportData));

            //The following runs the import as a foreground task, you need this when debugging
            //object stateInfo = null;
            //importThread.ImportData(stateInfo);
        }

        /// <summary>
        /// Updates the form state when the import has finished
        /// </summary>
        public void ProgressForm_OnFinish()
        {
            //Change the cursor back to the default
            this.Cursor = Cursors.Default;

            //Enable the Exit button and disable cancel
            this.btnCancel.Enabled = false;
            this.btnExit.Enabled = true;

            //Display a message
            MessageBox.Show("SpiraTest Import to TestLink Successful!", "Import", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// <summary>
        /// Displays any errors raised by the import thread process
        /// </summary>
        /// <param name="exception">The exception raised</param>
        public void ProgressForm_OnError(Exception exception)
        {
            //Change the cursor back to the default
            this.Cursor = Cursors.Default;

            //Enable the Exit button and disable cancel
            this.btnCancel.Enabled = false;
            this.btnExit.Enabled = true;

            //Display the exception error message
            MessageBox.Show("SpiraTest Import Failed. Error: " + exception.Message, "Import", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        /// <summary>
        /// Displays any errors raised by the import thread process
        /// </summary>
        /// <param name="exception">The exception raised</param>
        public void ProgressForm_OnError(string message)
        {
            //Change the cursor back to the default
            this.Cursor = Cursors.Default;

            //Enable the Exit button and disable cancel
            this.btnCancel.Enabled = false;
            this.btnExit.Enabled = true;

            //Display the exception error message
            MessageBox.Show(message, "Import", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        /// <summary>
        /// Updates the progress display of the form
        /// </summary>
        /// <param name="progress">The progress state</param>
        public void ProgressForm_OnProgressUpdate(int progress)
        {
            //Make all the controls up to the specified one visible
            for (int i = 1; i <= 6; i++)
            {
                this.Controls.Find("lblProgress" + i.ToString(), true)[0].Visible = (i <= progress);
            }

            //Also update the progress bar
            this.progressBar1.Value = progress;
        }

        /*
        /// <summary>
        /// Imports a single custom field value
        /// </summary>
        /// <param name="qcObject"></param>
        /// <param name="remoteArtifact"></param>
        private void ImportCustomField(StreamWriter streamWriter, IBaseField qcObject, string qcFieldName, List<RemoteArtifactCustomProperty> artifactCustomProperties)
        {
            try
            {
                string qcFieldNameUpperCase = qcFieldName.ToUpperInvariant();
                if (qcObject[qcFieldName] != null && qcObject[qcFieldName] is String)
                {
                    string fieldValue = (string)qcObject[qcFieldName];

                    //Find the matching custom property/type
                    if (this.customPropertyMapping.ContainsKey(qcFieldNameUpperCase))
                    {
                        int propertyNumber = this.customPropertyMapping[qcFieldNameUpperCase];
                        if (this.customPropertyMappingType.ContainsKey(qcFieldNameUpperCase))
                        {
                            int customPropertyTypeId = this.customPropertyMappingType[qcFieldNameUpperCase];

                            //Handle each type separately
                            switch ((Utils.CustomPropertyTypeEnum)customPropertyTypeId)
                            {
                                case Utils.CustomPropertyTypeEnum.Text:
                                    {
                                        RemoteArtifactCustomProperty remoteArtifactCustomProperty = new RemoteArtifactCustomProperty();
                                        artifactCustomProperties.Add(remoteArtifactCustomProperty);
                                        remoteArtifactCustomProperty.PropertyNumber = propertyNumber;
                                        remoteArtifactCustomProperty.StringValue = fieldValue;
                                    }
                                    break;

                                case Utils.CustomPropertyTypeEnum.List:
                                    {
                                        //Get the custom property value id                                    
                                        if (this.customValueNameMapping.ContainsKey(fieldValue))
                                        {
                                            int customPropertyValueId = this.customValueNameMapping[fieldValue];
                                            RemoteArtifactCustomProperty remoteArtifactCustomProperty = new RemoteArtifactCustomProperty();
                                            artifactCustomProperties.Add(remoteArtifactCustomProperty);
                                            remoteArtifactCustomProperty.PropertyNumber = propertyNumber;
                                            remoteArtifactCustomProperty.IntegerValue = customPropertyValueId;
                                        }
                                        else
                                        {
                                            streamWriter.WriteLine(String.Format("Warning: Unable top find a matching custom property value for field value '{0}', so leaving custom property '{1}' unset.", fieldValue, qcFieldName));
                                        }
                                    }
                                    break;

                                case Utils.CustomPropertyTypeEnum.MultiList:
                                    {
                                        //Split the custom values (by semicolon)
                                        string[] values = fieldValue.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
                                        if (values.Length > 0)
                                        {
                                            List<int> customPropertyValueIds = new List<int>();
                                            foreach (string customFieldValue in values)
                                            {
                                                //Get the custom property value id                                    
                                                if (this.customValueNameMapping.ContainsKey(customFieldValue))
                                                {
                                                    int customPropertyValueId = this.customValueNameMapping[customFieldValue];
                                                    customPropertyValueIds.Add(customPropertyValueId);
                                                }
                                                else
                                                {
                                                    streamWriter.WriteLine(String.Format("Warning: Unable top find a matching custom property value for field value '{0}', so leaving custom property '{1}' unset.", fieldValue, qcFieldName));
                                                }
                                            }
                                            if (customPropertyValueIds.Count > 0)
                                            {
                                                RemoteArtifactCustomProperty remoteArtifactCustomProperty = new RemoteArtifactCustomProperty();
                                                artifactCustomProperties.Add(remoteArtifactCustomProperty);
                                                remoteArtifactCustomProperty.PropertyNumber = propertyNumber;
                                                remoteArtifactCustomProperty.IntegerListValue = customPropertyValueIds.ToArray();
                                            }
                                        }
                                    }
                                    break;

                                case Utils.CustomPropertyTypeEnum.User:
                                    {
                                        //Get the user id for the login
                                        string userLogin = fieldValue;
                                        if (this.usersMapping.ContainsKey(userLogin))
                                        {
                                            int userId = this.usersMapping[userLogin];
                                            RemoteArtifactCustomProperty remoteArtifactCustomProperty = new RemoteArtifactCustomProperty();
                                            artifactCustomProperties.Add(remoteArtifactCustomProperty);
                                            remoteArtifactCustomProperty.PropertyNumber = propertyNumber;
                                            remoteArtifactCustomProperty.IntegerValue = userId;
                                        }
                                    }
                                    break;
                            }
                        }
                        else
                        {
                            streamWriter.WriteLine("Warning: Unable to find TYPE mapping for QC user field '" + qcFieldName + "' for HP ALM object " + qcObject.ID + " so ignoring this custom field and continuing.");
                        }
                    }
                    else
                    {
                        streamWriter.WriteLine("Warning: Unable to find NAME mapping for QC user field '" + qcFieldName + "' for HP ALM object " + qcObject.ID + " so ignoring this custom field and continuing.");
                    }
                }
            }
            catch (Exception exception)
            {
                //If we have an error, log it and continue
                streamWriter.WriteLine("Warning: Unable to access QC user field '" + qcFieldName + "' for HP ALM object " + qcObject.ID + " so ignoring this custom field and continuing (" + exception.Message + ")");
            }
        }*/

            /*
        private int ImportProject(StreamWriter streamWriter)
        {
            //Get the basic data from the OTA API
            MainFormHandle.EnsureConnected();
            string name = MainFormHandle.TdConnection.ProjectName;
            RemoteProject remoteProject = new RemoteProject();
            remoteProject.Name = name;
            remoteProject.Description = name;
            remoteProject.Active = true;

            //Use the SiteAdmin api to get the description, if the user doesn't have permissions, log and carry on
            try
            {
                //Get the project meta-data
                string projectXml = MainFormHandle.SaConnection.GetProject(MainFormHandle.TdConnection.DomainName, MainFormHandle.TdConnection.ProjectName);
                //streamWriter.WriteLine("Project XML: " + projectXml);
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.LoadXml(projectXml);
                XmlNode xmlDescNode = xmlDoc.SelectSingleNode("/TDXItem/DESCRIPTION");
                if (xmlDescNode != null)
                {
                    remoteProject.Description = MakeXmlSafe(xmlDescNode.InnerText);
                }
            }
            catch (Exception exception)
            {
                streamWriter.WriteLine("Unable to access the project through the SiteAdmin API, so project description will not be imported: " + exception.Message);
            }
            
            remoteProject = ImportFormHandle.SpiraImportProxy.Project_Create(remoteProject, null);
            streamWriter.WriteLine("New Project '" + name + "' Created");
            int projectId = remoteProject.ProjectId.Value;

            //Now we need to get the custom lists for the project
            Dictionary<int, int> customListMapping = new Dictionary<int, int>();
            Customization customization = MainFormHandle.TdConnection.Customization;

            CustomizationLists projectLists = customization.Lists;
            if (projectLists != null)
            {
                //the lists are a hierarchy (ROOT>LIST>VALUE), so get the root node
                for (int i = 1; i <= projectLists.Count; i++)
                {
                    CustomizationList projectList = projectLists.get_ListByCount(i);
                    string listName = projectList.Name;
                    bool isListDeleted = projectList.Deleted;

                    streamWriter.WriteLine(String.Format("Adding Project List: {0} to Spira", listName));

                    //Create the Spira custom list
                    RemoteCustomList remoteCustomList = new RemoteCustomList();
                    remoteCustomList.Name = listName;
                    remoteCustomList.Active = !isListDeleted;
                    int newCustomListId = ImportFormHandle.SpiraImportProxy.CustomProperty_AddCustomList(remoteCustomList).CustomPropertyListId.Value;
                    CustomizationListNode rootNode = projectList.RootNode;
                    if (!customListMapping.ContainsKey(rootNode.ID))
                    {
                        customListMapping.Add(rootNode.ID, newCustomListId);
                    }

                    //Get the children (values)
                    foreach (CustomizationListNode valueNode in rootNode.Children)
                    {
                        int valueNodeId = valueNode.ID;
                        string valueNodeName = valueNode.Name;
                        bool isValueDeleted = valueNode.Deleted;
                        string type = valueNode.Type;

                        //Create the Spira custom list values
                        RemoteCustomListValue remoteCustomListValue = new RemoteCustomListValue();
                        remoteCustomListValue.CustomPropertyListId = newCustomListId;
                        remoteCustomListValue.Name = valueNodeName;
                        //IsDeleted and IsActive fields may be added in future versions of the system
                        int newCustomListValueId = ImportFormHandle.SpiraImportProxy.CustomProperty_AddCustomListValue(remoteCustomListValue).CustomPropertyValueId.Value;

                        //Add to mapping (we map both name and ID)
                        streamWriter.WriteLine(String.Format("Adding Project List Value: {0} (ID={1}) to Spira, new Spira ID={2}", valueNodeName, valueNodeId, newCustomListValueId));
                        if (!this.customValueMapping.ContainsKey(valueNodeId))
                        {
                            this.customValueMapping.Add(valueNodeId, newCustomListValueId);
                        }
                        if (!this.customValueNameMapping.ContainsKey(valueNodeName))
                        {
                            this.customValueNameMapping.Add(valueNodeName, newCustomListValueId);
                        }
                    }
                }
            }

            //Now we need to get the custom fields defined in the project
            CustomizationFields customFieldDefinitions = (CustomizationFields)customization.Fields;
            if (customFieldDefinitions != null)
            {
                List customFields = customFieldDefinitions.get_Fields();
                Dictionary<int, int> artifactPropertyNumber = new Dictionary<int, int>();
                foreach (CustomizationField customField in customFields)
                {
                    //We only want custom fields
                    if (!customField.IsSystem && customField.IsActive)
                    {
                        string displayName = customField.UserLabel;
                        string fieldName = customField.ColumnName;
                        bool isMultiValue = customField.IsMultiValue;
                        bool isRequired = customField.IsRequired;   //We don't import options so not currently used
                        string tableName = customField.TableName;
                        string editType = customField.UserColumnType;

                        //streamWriter.WriteLine("Debug3: " + customField.RootId + "," + customField.ColumnType + "," + customField.UserColumnType + "," + customField.EditMask);

                        //See if we have a matching custom list
                        int? customPropertyListId = null;
                        if (customField.RootId != null)
                        {
                            int listRootId = customField.RootId;
                            if (customListMapping.ContainsKey(listRootId))
                            {
                                customPropertyListId = customListMapping[listRootId];
                            }
                        }

                        //Get the corresponding artifact id and custom property type
                        int? artifactTypeId = ConvertArtifactType(tableName);
                        int customPropertyTypeId = ConvertCustomPropertyType(customPropertyListId.HasValue, isMultiValue);

                        if (artifactTypeId.HasValue)
                        {
                            //Get the next available property number for this artifact type
                            int propertyNumber = 1;
                            if (artifactPropertyNumber.ContainsKey(artifactTypeId.Value))
                            {
                                propertyNumber = artifactPropertyNumber[artifactTypeId.Value] + 1;
                                artifactPropertyNumber[artifactTypeId.Value] = propertyNumber;
                            }
                            else
                            {
                                artifactPropertyNumber.Add(artifactTypeId.Value, propertyNumber);
                            }

                            //Add to Spira
                            RemoteCustomProperty remoteCustomProperty = new RemoteCustomProperty();
                            remoteCustomProperty.Name = displayName;
                            remoteCustomProperty.PropertyNumber = propertyNumber;
                            remoteCustomProperty.ProjectId = projectId;
                            remoteCustomProperty.CustomPropertyTypeId = customPropertyTypeId;
                            remoteCustomProperty.ArtifactTypeId = artifactTypeId.Value;
                            int customPropertyId = ImportFormHandle.SpiraImportProxy.CustomProperty_AddDefinition(remoteCustomProperty, customPropertyListId).CustomPropertyId.Value;

                            //Add to the mapping dictionaries
                            if (!this.customPropertyMapping.ContainsKey(fieldName))
                            {
                                this.customPropertyMapping.Add(fieldName, propertyNumber);
                            }
                            if (!this.customPropertyMappingType.ContainsKey(fieldName))
                            {
                                this.customPropertyMappingType.Add(fieldName, customPropertyTypeId);
                                streamWriter.WriteLine(String.Format("Adding Custom Property Mapping for: {0}=Custom_{1}:Type={2}:Caption={3}", fieldName, propertyNumber, customPropertyTypeId, displayName));
                            }
                        }
                        else
                        {
                            streamWriter.WriteLine(String.Format("WARNING: Unable to find artifact type for system table '{0}'", tableName));
                        }
                    }
                }

                //If we have at least one custom field type remaining for defects, add a special field to capture the legacy defect ID
                if (artifactPropertyNumber.ContainsKey((int)Utils.ArtifactType.Incident))
                {
                    int propertyNumber = artifactPropertyNumber[(int)Utils.ArtifactType.Incident];
                    if (propertyNumber < 30)
                    {
                        //Add to Spira
                        RemoteCustomProperty remoteCustomProperty = new RemoteCustomProperty();
                        remoteCustomProperty.Name = "ALM Defect #";
                        remoteCustomProperty.PropertyNumber = propertyNumber + 1;
                        remoteCustomProperty.ProjectId = projectId;
                        remoteCustomProperty.CustomPropertyTypeId = (int)Utils.CustomPropertyTypeEnum.Integer;
                        remoteCustomProperty.ArtifactTypeId = (int)Utils.ArtifactType.Incident;
                        int customPropertyId = ImportFormHandle.SpiraImportProxy.CustomProperty_AddDefinition(remoteCustomProperty, null).CustomPropertyId.Value;

                        //Add to mapping
                        if (!this.customPropertyMapping.ContainsKey(Utils.CUSTOM_FIELD_ALM_DEFECT_ID))
                        {
                            this.customPropertyMapping.Add(Utils.CUSTOM_FIELD_ALM_DEFECT_ID, propertyNumber);
                        }
                    }
                }
            }

            return remoteProject.ProjectId.Value;
        }

        /// <summary>
        /// Imports a list of attachments for an artifact from QC extended storage
        /// </summary>
        /// <param name="streamWriter">The streamwriter used to log messages</param>
        /// <param name="entityType">The QC entity name</param>
        /// <param name="entityId">The QC entity ID</param>
        /// <param name="attachmentFactory">The QC attachment factory object</param>
        /// <param name="artifactId">The Spira artifact id</param>
        /// <param name="artifactTypeId">The Spira artifact type id</param>
        private void ImportAttachments(StreamWriter streamWriter, string entityType, int entityId, object attachmentFactory, int artifactId, Utils.ArtifactType artifactType)
        {
            //Make sure we have a populated attachment factory
            if (attachmentFactory == null)
            {
                return;
            }

            if (attachmentFactory is HP.ALM.AttachmentFactory)
            {
                try
                {
                    //Create a temporary folder to download the attachments to
                    string localSettings = System.Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                    if (!Directory.Exists(localSettings + "\\Inflectra"))
                    {
                        Directory.CreateDirectory(localSettings + "\\Inflectra");
                    }
                    if (!Directory.Exists(localSettings + "\\Inflectra\\QCImport"))
                    {
                        Directory.CreateDirectory(localSettings + "\\Inflectra\\QCImport");
                    }

                    //Now get the list of attachments and import
                    streamWriter.WriteLine("Retrieving Attachments for " + entityType + " '" + entityId + "'");
                    HP.ALM.List attachmentFactoryList = ((HP.ALM.AttachmentFactory)attachmentFactory).NewList("");
                    foreach (HP.ALM.Attachment attachment in attachmentFactoryList)
                    {
                        string displayFilename = (string)attachment.get_Name(1);
                        string filename = (string)attachment.get_Name(0);
                        streamWriter.WriteLine("Retrieving Attachment '" + filename + "'");

                        string description = "";
                        if (attachment.Description != null)
                        {
                            description = MakeXmlSafe(attachment.Description);
                        }

                        //Now we need to actually get the attachment binary data
                        HP.ALM.IExtendedStorage attachmentStorage = (HP.ALM.IExtendedStorage)attachment.AttachmentStorage;
                        attachmentStorage.ClientPath = localSettings + "\\Inflectra\\QCImport\\";
                        attachmentStorage.Load(filename, true);

                        //Now read the file back in and import
                        if (File.Exists(localSettings + "\\Inflectra\\QCImport\\" + filename))
                        {
                            FileStream fileStream = File.OpenRead(localSettings + "\\Inflectra\\QCImport\\" + filename);
                            int numBytes = (int)fileStream.Length;
                            byte[] attachmentData = new byte[numBytes];
                            fileStream.Read(attachmentData, 0, numBytes);
                            fileStream.Close();

                            //Import the attachment
                            RemoteDocument remoteDocument = new RemoteDocument();
                            remoteDocument.FilenameOrUrl = displayFilename;
                            remoteDocument.Description = MakeXmlSafe(description);
                            remoteDocument.ArtifactId = artifactId;
                            remoteDocument.ArtifactTypeId = (int)artifactType;
                            ImportFormHandle.SpiraImportProxy.Document_AddFile(remoteDocument, attachmentData);
                        }
                        else
                        {
                            streamWriter.WriteLine("Unable to import attachment '" + filename + "' - file data not readable from QualityCenter.");
                        }
                    }
                }
                catch (Exception exception)
                {
                    //Log and continue
                    streamWriter.WriteLine("Warning: Unable to import attachments for entity id '" + entityId + "' - continuing with import (" + exception.Message + ").");
                }
            }
        }*/

        /// <summary>
        /// Replaces the QualityCenter parameter token with the SpiraTest equivalent.
        /// </summary>
        /// <param name="input">The input string</param>
        /// <returns>The output string</returns>
        /// <remarks>
        /// Quality center uses <<<Parameter Name>>> and SpiraTest uses
        ///	${parameter name} so we need to account for the potential
        ///	case difference as well. Also need to handle HTML encoded version as well
        /// </remarks>
        private string ReplaceParameterToken(string input)
        {
            string output = Regex.Replace(input, @"&lt;&lt;&lt;[a-zA-Z0-9 /:\-_\.]+&gt;&gt;&gt;", new MatchEvaluator(MakeParameterToken));
            output = Regex.Replace(output, @"\<\<\<[a-zA-Z0-9 /:\-_\.]+\>\>\>", new MatchEvaluator(MakeParameterToken2));
            return output;
        }

        private string MakeParameterToken(Match m)
        {
            // Get the matched string.
            string x = m.ToString();

            //Replace the <<< with ${ and >>> with }
            x = x.Replace("&lt;&lt;&lt;", "${");
            x = x.Replace("&gt;&gt;&gt;", "}");

            //Remove any spaces and make them underscores
            x = x.Replace(" ", "_");
            return x.ToLower();
        }
        private string MakeParameterToken2(Match m)
        {
            // Get the matched string.
            string x = m.ToString();

            //Replace the <<< with ${ and >>> with }
            x = x.Replace("<<<", "${");
            x = x.Replace(">>>", "}");

            //Remove any spaces and make them underscores
            x = x.Replace(" ", "_");
            return x.ToLower();
        }

        /// <summary>
        /// Converts a QualityCenter test set status for use in SpiraTest
        /// </summary>
        /// <param name="status">The TD execution status</param>
        /// <returns>The SpiraTest execution status</returns>
        protected int ConvertTestSetStatus(string status)
        {
            int statusId = 1; //Default to not started (no nulls allowed)
            switch (status)
            {
                case "Open":
                    //Not Started
                    statusId = 1;
                    break;

                case "Closed":
                    //Deferred
                    statusId = 5;
                    break;

            }
            return statusId;
        }

        /// <summary>
        /// Returns the Spira custom property type for the appropriate QC Edit Style (text if not one of the known ones)
        /// </summary>
        /// <param name="editType"></param>
        /// <returns></returns>
        private int ConvertCustomPropertyType(bool listProvided, bool isMultiValue)
        {
            int customPropertyTypeId = (int)Utils.CustomPropertyTypeEnum.Text;
            /*  If they every allow you to access the SF_EDIT_TYPE field through the API, could get a better type relationshop
            switch (editType)
            {
                case "ListCombo":
                case "TreeCombo":
                    {
                        if (isMultiValue)
                        {
                            customPropertyTypeId = (int)Utils.CustomPropertyTypeEnum.MultiList;
                        }
                        else
                        {
                            customPropertyTypeId = (int)Utils.CustomPropertyTypeEnum.List;
                        }
                    }
                    break;

                case "UserCombo":
                    customPropertyTypeId = (int)Utils.CustomPropertyTypeEnum.User;
                    break;
            }*/
            if (listProvided)
            {
                if (isMultiValue)
                {
                    customPropertyTypeId = (int)Utils.CustomPropertyTypeEnum.MultiList;
                }
                else
                {
                    customPropertyTypeId = (int)Utils.CustomPropertyTypeEnum.List;
                }
            }

            return customPropertyTypeId;
        }

        /*
        /// <summary>
        /// Converts a QC artifact type into the corresponding Spira artifact type
        /// </summary>
        /// <param name="qcType"></param>
        /// <returns></returns>
        private int? ConvertArtifactType(string qcType)
        {
            int? artifactTypeId = null;
            switch (qcType)
            {
                case "REQ":
                    artifactTypeId = (int)Utils.ArtifactType.Requirement;
                    break;

                case "RELEASES":
                    artifactTypeId = (int)Utils.ArtifactType.Release;
                    break;

                case "TEST":
                    artifactTypeId = (int)Utils.ArtifactType.TestCase;
                    break;

                case "RUN":
                    artifactTypeId = (int)Utils.ArtifactType.TestRun;
                    break;

                case "BUG":
                    artifactTypeId = (int)Utils.ArtifactType.Incident;
                    break;

                case "CYCLE":
                    artifactTypeId = (int)Utils.ArtifactType.TestSet;
                    break;

                case "DESSTEP":
                    artifactTypeId = (int)Utils.ArtifactType.TestStep;
                    break;
            }

            return artifactTypeId;
        }

        /// <summary>
        /// Converts a QualityCenter execution status for use in SpiraTest
        /// </summary>
        /// <param name="status">The TD execution status</param>
        /// <returns>The SpiraTest execution status</returns>
        protected int ConvertExecutionStatus(string status)
        {
            int statusId = 3; //Default to not run (no nulls allowed)
            switch (status)
            {
                case "Failed":
                    //Failed
                    statusId = 1;
                    break;
                case "Passed":
                    //Passed
                    statusId = 2;
                    break;
                case "N/A":
                    //Not Run
                    statusId = 3;
                    break;
                case "No Run":
                    //Not Run
                    statusId = 3;
                    break;
                case "Not Completed":
                    //Not Run
                    statusId = 3;
                    break;
            }
            return statusId;
        }

        /// <summary>
        /// Makes a string safe for use in XML (e.g. web service)
        /// </summary>
        /// <param name="input">The input string (as object)</param>
        /// <returns>The output string</returns>
        protected string MakeXmlSafe(object input)
        {
            //Handle null reference case
            if (input == null)
            {
                return "";
            }

            //Handle the case where the object is not a string
            string inputString;
            if (input.GetType() == typeof(string))
            {
                inputString = (string)input;
            }
            else
            {
                inputString = input.ToString();
            }

            //Handle empty string case
            if (inputString == "")
            {
                return inputString;
            }

            string output = inputString.Replace("\x00", "");
            output = output.Replace("\x01", "");
            output = output.Replace("\x02", "");
            output = output.Replace("\x03", "");
            output = output.Replace("\x04", "");
            output = output.Replace("\x05", "");
            output = output.Replace("\x06", "");
            output = output.Replace("\x07", "");
            output = output.Replace("\x08", "");
            output = output.Replace("\x0B", "");
            output = output.Replace("\x0C", "");
            output = output.Replace("\x0E", "");
            output = output.Replace("\x0F", "");
            output = output.Replace("\x10", "");
            output = output.Replace("\x11", "");
            output = output.Replace("\x12", "");
            output = output.Replace("\x13", "");
            output = output.Replace("\x14", "");
            output = output.Replace("\x15", "");
            output = output.Replace("\x16", "");
            output = output.Replace("\x17", "");
            output = output.Replace("\x18", "");
            output = output.Replace("\x19", "");
            output = output.Replace("\x1A", "");
            output = output.Replace("\x1B", "");
            output = output.Replace("\x1C", "");
            output = output.Replace("\x1D", "");
            output = output.Replace("\x1E", "");
            output = output.Replace("\x1F", "");
            return output;
        }

        /// <summary>
        /// Converts a TestDirector priority for use in SpiraTest
        /// </summary>
        /// <param name="priority">The TD priority</param>
        /// <param name="spiraImport">Handle to the import API</param>
        /// <returns>The SpiraTest priority</returns>
        /// <remarks>Inserts the priority into SpiraTest upon first usage then maps for subsequent</remarks>
        protected Nullable<int> ConvertPriority(string priority, SpiraSoapService.SoapServiceClient spiraImport)
        {
            //Handle NULL case
            if (priority == null || priority == "")
            {
                return null;
            }

            //First see if we have the priority in our mapping hashtable
            if (this.incidentPriorityMapping.ContainsKey(priority))
            {
                return (int)this.incidentPriorityMapping[priority];
            }
            else
            {
                //Add a new custom incident priority to SpiraTest - default to color=silver
                RemoteIncidentPriority remoteIncidentPriority = new RemoteIncidentPriority();
                remoteIncidentPriority.Name = priority;
                remoteIncidentPriority.Active = true;
                remoteIncidentPriority.Color = "eeeeee";
                Nullable<int> priorityId = spiraImport.Incident_AddPriority(remoteIncidentPriority).PriorityId;
                if (priorityId.HasValue)
                {
                    this.incidentPriorityMapping.Add(priority, priorityId.Value);
                }
                return priorityId;
            }
        }

        /// <summary>
        /// Converts a TestDirector severity for use in SpiraTest
        /// </summary>
        /// <param name="severity">The TD severity</param>
        /// <param name="spiraImport">Handle to the import API</param>
        /// <returns>The SpiraTest severity</returns>
        /// <remarks>Inserts the severity into SpiraTest upon first usage then maps for subsequent</remarks>
        protected Nullable<int> ConvertSeverity(string severity, SpiraSoapService.SoapServiceClient spiraImport)
        {
            //Handle NULL case
            if (severity == null || severity == "")
            {
                return null;
            }

            //First see if we have the severity in our mapping hashtable
            if (this.incidentSeverityMapping.ContainsKey(severity))
            {
                return (int)this.incidentSeverityMapping[severity];
            }
            else
            {
                //Add a new custom incident severity to SpiraTest - default to color=silver
                RemoteIncidentSeverity remoteIncidentSeverity = new RemoteIncidentSeverity();
                remoteIncidentSeverity.Name = severity;
                remoteIncidentSeverity.Color = "eeeeee";
                remoteIncidentSeverity.Active = true;
                Nullable<int> severityId = spiraImport.Incident_AddSeverity(remoteIncidentSeverity).SeverityId;
                if (severityId.HasValue)
                {
                    this.incidentSeverityMapping.Add(severity, severityId.Value);
                }
                return severityId;
            }
        }

        /// <summary>
        /// Converts a TestDirector incident status for use in SpiraTest
        /// </summary>
        /// <param name="status">The TD status</param>
        /// <param name="spiraImport">Handle to the import API</param>
        /// <returns>The SpiraTest status</returns>
        /// <remarks>Inserts the status into SpiraTest upon first usage then maps for subsequent</remarks>
        protected Nullable<int> ConvertStatus(string status, SpiraSoapService.SoapServiceClient spiraImport)
        {
            //Handle NULL case
            if (status == null || status == "")
            {
                return null;
            }

            //First see if we have the status in our mapping hashtable
            if (this.incidentStatusMapping.ContainsKey(status))
            {
                return (int)this.incidentStatusMapping[status];
            }
            else
            {
                //Add a new custom incident status to SpiraTest
                RemoteIncidentStatus remoteIncidentStatus = new RemoteIncidentStatus();
                remoteIncidentStatus.Name = status;
                remoteIncidentStatus.Active = true;
                Nullable<int> incidentStatusId = spiraImport.Incident_AddStatus(remoteIncidentStatus).IncidentStatusId;
                if (incidentStatusId.HasValue)
                {
                    this.incidentStatusMapping.Add(status, incidentStatusId.Value);
                }
                return incidentStatusId;
            }
        }

        /// <summary>
        /// Converts a HP ALM requirement status for use in SpiraTest
        /// </summary>
        /// <param name="reviewStatus">The HP review status</param>
        /// <returns>The SpiraTest importance</returns>
        protected int ConvertReqStatus(string reviewStatus)
        {
            int statusId = (int)Utils.RequirementStatusEnum.Requested;
            switch (reviewStatus)
            {
                case "Not Reviewed":
                    //Low
                    statusId = (int)Utils.RequirementStatusEnum.Requested;
                    break;
                case "Reviewed":
                    //Medium
                    statusId = (int)Utils.RequirementStatusEnum.Accepted;
                    break;
                case "Changed":
                    //High
                    statusId = (int)Utils.RequirementStatusEnum.UnderReview;
                    break;
            }
            return statusId;
        }

        /// <summary>
        /// Converts a TestDirector requirement importance for use in SpiraTest
        /// </summary>
        /// <param name="importance">The TD importance</param>
        /// <returns>The SpiraTest importance</returns>
        protected Nullable<int> ConvertImportance(string importance)
        {
            Nullable<int> importanceId = null;
            switch (importance)
            {
                case "1-Low":
                    //Low
                    importanceId = (int)Utils.RequirementImportanceEnum.Low;
                    break;
                case "2-Medium":
                    //Medium
                    importanceId = (int)Utils.RequirementImportanceEnum.Medium;
                    break;
                case "3-High":
                    //High
                    importanceId = (int)Utils.RequirementImportanceEnum.High;
                    break;
                case "4-Very High":
                    //Very High
                    importanceId = (int)Utils.RequirementImportanceEnum.Critical;
                    break;
                case "5-Urgent":
                    //Urgent
                    importanceId = (int)Utils.RequirementImportanceEnum.Critical;
                    break;
            }
            return importanceId;
        }

        /// <summary>
        /// Retrieves the SpiraTest test run step id from a linked Quality Center bug id
        /// </summary>
        /// <param name="bugObject">The QC bug object</param>
        /// <returns>The test run step of the test step that generated the bug</returns>
        protected Nullable<int> GetTestRunStepIdForBug(Bug bugObject)
        {
            Nullable<int> testRunStepId = null;

            //First get the qc test step id for this bug
            if (qcBugToRunStepMapping.ContainsKey(bugObject.ID))
            {
                int runStepId = qcBugToRunStepMapping[bugObject.ID];
                //Lookup the mapping to the Spira test run step ID
                if (this.testRunStepMapping.ContainsKey(runStepId))
                {
                    testRunStepId = (int)this.testRunStepMapping[runStepId];
                }
            }
            return testRunStepId;
        }

        /// <summary>
        /// Imports release folders into Spira as 'parent' releases
        /// </summary>
        /// <param name="streamWriter">The logfile writer</param>
        /// <param name="releaseFactory">The QC release folder factory object</param>
        /// <param name="parentReleaseId">The parent Spira release, or null for top-level</param>
        protected void ImportReleaseFolders(StreamWriter streamWriter, HP.ALM.ReleaseFolderFactory relFolderFactory, int? parentReleaseId)
        {
            HP.ALM.TDFilter tdFilter = (HP.ALM.TDFilter)relFolderFactory.Filter;
            tdFilter.set_Order("rf_path", 1);
            HP.ALM.List relFolderList;
            try
            {
                relFolderList = (HP.ALM.List)tdFilter.NewList();
            }
            catch (Exception)
            {
                //If we get an exception, try to reconnect and re-execute
                MainFormHandle.TryReconnect();
                relFolderFactory = (HP.ALM.ReleaseFolderFactory)MainFormHandle.TdConnection.ReleaseFolderFactory;
                tdFilter = (HP.ALM.TDFilter)relFolderFactory.Filter;
                relFolderList = (HP.ALM.List)tdFilter.NewList();
            }

            //Import the release folders as top-level releases in Spira
            foreach (HP.ALM.ReleaseFolder relFolderObject in relFolderList)
            {
                //Extract the release folder info
                int releaseFolderId = (int)relFolderObject.ID;

                //The factory object returns all children, not just the direct ones, so make sure we only load a release once
                if (!releaseFolderMapping.ContainsKey(releaseFolderId))
                {
                    //Load the release and capture the ID
                    RemoteRelease remoteRelease = new RemoteRelease();
                    remoteRelease.Name = relFolderObject.Name;
                    remoteRelease.VersionNumber = relFolderObject.Name.SafeSubstring(10);
                    if (relFolderObject["rf_description"] != null)
                    {
                        remoteRelease.Description = MakeXmlSafe(relFolderObject["rf_description"]);
                    }
                    //Dummy dates for folders
                    remoteRelease.StartDate = DateTime.UtcNow;
                    remoteRelease.EndDate = DateTime.UtcNow.AddMonths(1);
                    remoteRelease.ResourceCount = 1;
                    remoteRelease.Active = true;
                    int newReleaseId = ImportFormHandle.SpiraImportProxy.Release_Create(remoteRelease, parentReleaseId).ReleaseId.Value;
                    streamWriter.WriteLine("Added release folder: '" + releaseFolderId);

                    //Add attachments if requested
                    if (this.MainFormHandle.chkImportAttachments.Checked)
                    {
                        try
                        {
                            ImportAttachments(streamWriter, "RELEASE_FOLDERS", releaseFolderId, relFolderObject.Attachments, newReleaseId, Utils.ArtifactType.Release);
                        }
                        catch (Exception exception)
                        {
                            streamWriter.WriteLine("Warning: Unable to import attachments for release folder " + releaseFolderId + " (" + exception.Message + ")");
                        }
                    }

                    //Add to the mapping hashtables
                    if (!this.releaseFolderMapping.ContainsKey(releaseFolderId))
                    {
                        this.releaseFolderMapping.Add(releaseFolderId, newReleaseId);
                    }

                    //Import any child release folders
                    ImportReleaseFolders(streamWriter, relFolderObject.ReleaseFolderFactory, newReleaseId);

                    //Now import the child releases
                    ImportReleases(streamWriter, relFolderObject.ReleaseFactory, newReleaseId);
                }
            }
        }*/

        /*
        /// <summary>
        /// Imports automation hosts into Spira
        /// </summary>
        /// <param name="streamWriter">The logfile writer</param>
        /// <param name="hostFactory">The QC host factory object</param>
        protected void ImportHosts(StreamWriter streamWriter, HP.ALM.HostFactory hostFactory)
        {
            HP.ALM.TDFilter tdFilter = (HP.ALM.TDFilter)hostFactory.Filter;
            tdFilter.set_Order("ho_name", 1);
            HP.ALM.List hostList = (HP.ALM.List)tdFilter.NewList();

            //Import the hosts into Spira
            foreach (HP.ALM.Host hostObject in hostList)
            {
                //Extract the host info
                string hostToken = (string)hostObject.ID;

                //Load the automation host and capture the ID
                RemoteAutomationHost remoteAutomationHost = new RemoteAutomationHost();
                remoteAutomationHost.Name = hostObject.Name;
                remoteAutomationHost.Token = hostToken;
                remoteAutomationHost.Active = true;
                remoteAutomationHost.Description = MakeXmlSafe(hostObject.Description);
                int newAutomationHostId = ImportFormHandle.SpiraImportProxy.AutomationHost_Create(remoteAutomationHost).AutomationHostId.Value;
                streamWriter.WriteLine("Added host: '" + hostToken);

                //Add to the mapping hashtables
                if (!this.hostMapping.ContainsKey(hostToken))
                {
                    this.hostMapping.Add(hostToken, newAutomationHostId);
                }
            }
        }

        protected void ImportRequirements(StreamWriter streamWriter, HP.ALM.ReqFactory reqFactory)
        {
            HP.ALM.TDFilter tdFilter = (HP.ALM.TDFilter)reqFactory.Filter;
            tdFilter.set_Order("rq_req_path", 1);
            HP.ALM.List reqList;
            try
            {
                reqList = (HP.ALM.List)tdFilter.NewList();
            }
            catch (Exception)
            {
                //If we get an exception, try to reconnect and re-execute
                MainFormHandle.TryReconnect();
                reqFactory = (HP.ALM.ReqFactory)MainFormHandle.TdConnection.ReqFactory;
                tdFilter = (HP.ALM.TDFilter)reqFactory.Filter;
                reqList = (HP.ALM.List)tdFilter.NewList();
            }

            string lastRequirementPath = "";
            foreach (HP.ALM.Req reqObject in reqList)
            {
                //Extract the requirement info
                int requirementId = (int)reqObject.ID;
                string requirementPath = "";
                try
                {
                    requirementPath = reqObject["rq_req_path"];
                }
                catch (Exception exception)
                {
                    //Folders sometimes don't have paths
                    streamWriter.WriteLine(String.Format("Warning: unable to access rq_req_path for requirement {0} so ignoring path! ({1})", requirementId, exception.Message));
                }
                streamWriter.WriteLine(String.Format("Importing Requirement: {0} with path '{1}'", requirementId, requirementPath));
                string requirementPriority = reqObject.Priority;
                string name = MakeXmlSafe(reqObject.Name);
                string requirementComments = "";
                if (!String.IsNullOrEmpty(reqObject.Comment))
                {
                    //TD comments will be used for SpiraTest's description field
                    requirementComments = MakeXmlSafe(reqObject.Comment);
                }
                string authorName = reqObject.Author;

                //Convert the priority (TD has 5 states, SpiraTest has 4)
                Nullable<int> importanceId = ConvertImportance(requirementPriority);

                //Convert the status
                string reqReviewStatus = reqObject.Reviewed;
                int statusId = ConvertReqStatus(reqReviewStatus);

                //Convert the path into a relative position
                int indentPosition = 0;
                if (lastRequirementPath != "")
                {
                    //Check to see if the current path is a different length to the last one
                    if (lastRequirementPath.Length != requirementPath.Length)
                    {
                        //Get the relative offset
                        indentPosition = (requirementPath.Length - lastRequirementPath.Length) / 3;
                    }
                }
                lastRequirementPath = requirementPath;

                //See if we're a folder (doesn't have all the fields)
                bool isFolder = false;
                try
                {
                    if (reqObject.Type != null && reqObject.Type.ToLowerInvariant() == "folder")
                    {
                        isFolder = true;
                    }
                }
                catch (Exception)
                {
                    //Folders sometimes throw an exception when you try and access the field
                    isFolder = true;
                }

                //See if we have a release specified
                int? releaseId = null;
                try
                {
                    if (!isFolder && reqObject["rq_target_rel"] != null)
                    {
                        int qcReleaseId = (int)reqObject["rq_target_rel"];
                        if (this.releaseMapping.ContainsKey(qcReleaseId))
                        {
                            releaseId = this.releaseMapping[qcReleaseId];
                            streamWriter.WriteLine("Found requirement release: RL" + releaseId.Value.ToString());
                            break;
                        }
                    }
                    else if (!isFolder && reqObject["rq_target_rel_varchar"] != null)
                    {
                        //They are semicolon-separated
                        string relList = (string)reqObject["rq_target_rel_varchar"];
                        //streamWriter.WriteLine("DEBUG1: " + relList);
                        string[] releases = relList.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
                        foreach (string qcRelease in releases)
                        {
                            int qcReleaseId;
                            if (Int32.TryParse(qcRelease, out qcReleaseId))
                            {
                                if (this.releaseMapping.ContainsKey(qcReleaseId))
                                {
                                    releaseId = this.releaseMapping[qcReleaseId];
                                    streamWriter.WriteLine("Found requirement release: RL" + releaseId.Value.ToString());
                                    break;
                                }
                            }
                        }
                    }
                }
                catch (Exception)
                {
                    //Some types don't let you access this field
                    streamWriter.WriteLine(String.Format("Requirement {0} of type {1} does not allow access to the Release field so ignoring.", reqObject.ID, reqObject.Type));
                }

                //See if we have an iteration specified
                int? iterationId = null;
                try
                {
                    if (!isFolder && reqObject["rq_target_rcyc"] != null)
                    {
                        int qcCycleId = (int)reqObject["rq_target_rcyc"];
                        if (this.iterationMapping.ContainsKey(qcCycleId))
                        {
                            iterationId = this.iterationMapping[qcCycleId];
                            streamWriter.WriteLine("Found requirement iteration: RL" + iterationId.Value.ToString());
                            break;
                        }
                    }
                    else if (!isFolder && reqObject["rq_target_rcyc_varchar"] != null)
                    {
                        //They are semicolon-separated
                        string cycList = (string)reqObject["rq_target_rcyc_varchar"];
                        //streamWriter.WriteLine("DEBUG1: " + cycList);
                        string[] cycles = cycList.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
                        foreach (string qcCycle in cycles)
                        {
                            int qcCycleId;
                            if (Int32.TryParse(qcCycle, out qcCycleId))
                            {
                                if (this.iterationMapping.ContainsKey(qcCycleId))
                                {
                                    iterationId = this.iterationMapping[qcCycleId];
                                    streamWriter.WriteLine("Found requirement iteration: RL" + iterationId.Value.ToString());
                                    break;
                                }
                            }
                        }
                    }
                }
                catch (Exception)
                {
                    //Some types don't let you access this field
                    streamWriter.WriteLine(String.Format("Requirement {0} of type {1} does not allow access to the Release field so ignoring.", reqObject.ID, reqObject.Type));
                }

                //If we have an iteration, use that instead of the release
                if (iterationId.HasValue)
                {
                    releaseId = iterationId.Value;
                }

                //Lookup the author from the user mapping hashtable
                Nullable<int> authorId = null;
                if (authorName != null && authorName != "")
                {
                    if (this.usersMapping.ContainsKey(authorName))
                    {
                        authorId = (int)this.usersMapping[authorName];
                    }
                }

                //Load the requirement and capture the new id
                int? newRequirementId = null;
                try
                {
                    //Populate the requirement
                    RemoteRequirement remoteRequirement = new RemoteRequirement();
                    remoteRequirement.StatusId = statusId;
                    remoteRequirement.ImportanceId = importanceId;
                    remoteRequirement.Name = name;
                    remoteRequirement.Description = MakeXmlSafe(requirementComments);
                    remoteRequirement.AuthorId = authorId;
                    remoteRequirement.ReleaseId = releaseId;

                    //Load any custom properties - QC stores them all as text
                    if (!isFolder)
                    {
                        try
                        {
                            List<RemoteArtifactCustomProperty> customProperties = new List<RemoteArtifactCustomProperty>();
                            ImportCustomField(streamWriter, reqObject, "rq_user_01", customProperties);
                            ImportCustomField(streamWriter, reqObject, "rq_user_02", customProperties);
                            ImportCustomField(streamWriter, reqObject, "rq_user_03", customProperties);
                            ImportCustomField(streamWriter, reqObject, "rq_user_04", customProperties);
                            ImportCustomField(streamWriter, reqObject, "rq_user_05", customProperties);
                            ImportCustomField(streamWriter, reqObject, "rq_user_06", customProperties);
                            ImportCustomField(streamWriter, reqObject, "rq_user_07", customProperties);
                            ImportCustomField(streamWriter, reqObject, "rq_user_08", customProperties);
                            ImportCustomField(streamWriter, reqObject, "rq_user_09", customProperties);
                            ImportCustomField(streamWriter, reqObject, "rq_user_10", customProperties);
                            ImportCustomField(streamWriter, reqObject, "rq_user_11", customProperties);
                            ImportCustomField(streamWriter, reqObject, "rq_user_12", customProperties);
                            ImportCustomField(streamWriter, reqObject, "rq_user_13", customProperties);
                            ImportCustomField(streamWriter, reqObject, "rq_user_14", customProperties);
                            ImportCustomField(streamWriter, reqObject, "rq_user_15", customProperties);
                            ImportCustomField(streamWriter, reqObject, "rq_user_16", customProperties);
                            ImportCustomField(streamWriter, reqObject, "rq_user_17", customProperties);
                            ImportCustomField(streamWriter, reqObject, "rq_user_18", customProperties);
                            ImportCustomField(streamWriter, reqObject, "rq_user_19", customProperties);
                            ImportCustomField(streamWriter, reqObject, "rq_user_20", customProperties);
                            ImportCustomField(streamWriter, reqObject, "rq_user_21", customProperties);
                            ImportCustomField(streamWriter, reqObject, "rq_user_22", customProperties);
                            ImportCustomField(streamWriter, reqObject, "rq_user_23", customProperties);
                            ImportCustomField(streamWriter, reqObject, "rq_user_24", customProperties);
                            remoteRequirement.CustomProperties = customProperties.ToArray();
                        }
                        catch (Exception exception)
                        {
                            //If we have an error, log it and continue
                            streamWriter.WriteLine("Warning: Unable to access QC user fields for requirement " + requirementId + " so ignoring custom fields and continuing (" + exception.Message + ")");
                        }
                    }
                    newRequirementId = ImportFormHandle.SpiraImportProxy.Requirement_Create1(remoteRequirement, indentPosition).RequirementId;
                }
                catch (Exception exception)
                {
                    //If we have an error, log it and continue
                    streamWriter.WriteLine("Error adding requirement " + requirementId + " to SpiraTest (" + exception.Message + ")");
                }
                if (newRequirementId.HasValue)
                {
                    //Add to the mapping hashtable
                    this.requirementsMapping.Add(requirementId, newRequirementId.Value);

                    //Add attachments if requested
                    if (this.MainFormHandle.chkImportAttachments.Checked)
                    {
                        try
                        {
                            ImportAttachments(streamWriter, "REQ", requirementId, reqObject.Attachments, newRequirementId.Value, Utils.ArtifactType.Requirement);
                        }
                        catch (Exception exception)
                        {
                            streamWriter.WriteLine("Warning: Unable to import attachments for requirement " + requirementId + " (" + exception.Message + ")");
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Imports releases into Spira
        /// </summary>
        /// <param name="streamWriter">The logfile writer</param>
        /// <param name="releaseFactory">The QC release factory object</param>
        /// <param name="parentReleaseId">The parent Spira release, or null for top-level</param>
        protected void ImportReleases(StreamWriter streamWriter, HP.ALM.ReleaseFactory releaseFactory, int? parentReleaseId)
        {
            HP.ALM.TDFilter tdFilter = (HP.ALM.TDFilter)releaseFactory.Filter;
            tdFilter.set_Order("rel_name", 1);
            HP.ALM.List releaseList = (HP.ALM.List)tdFilter.NewList();

            //Import the releases into Spira
            foreach (HP.ALM.Release releaseObject in releaseList)
            {
                //Extract the release info
                int releaseId = (int)releaseObject.ID;

                //The factory object returns all children, not just the direct ones, so make sure we only load a release once
                if (!releaseMapping.ContainsKey(releaseId))
                {
                    //Load the release and capture the ID
                    RemoteRelease remoteRelease = new RemoteRelease();
                    remoteRelease.Name = releaseObject.Name;
                    if (releaseObject["rel_description"] != null)
                    {
                        remoteRelease.Description = MakeXmlSafe(releaseObject["rel_description"]);
                    }
                    remoteRelease.VersionNumber = releaseObject.Name.SafeSubstring(10);
                    remoteRelease.StartDate = releaseObject.StartDate.ToUniversalTime();
                    remoteRelease.EndDate = releaseObject.EndDate.ToUniversalTime();
                    remoteRelease.ResourceCount = 1;
                    remoteRelease.Active = true;
                    int newReleaseId = ImportFormHandle.SpiraImportProxy.Release_Create(remoteRelease, parentReleaseId).ReleaseId.Value;
                    streamWriter.WriteLine("Added release: '" + releaseId);

                    //Add attachments if requested
                    if (this.MainFormHandle.chkImportAttachments.Checked)
                    {
                        try
                        {
                            ImportAttachments(streamWriter, "RELEASES", releaseId, releaseObject.Attachments, newReleaseId, Utils.ArtifactType.Release);
                        }
                        catch (Exception exception)
                        {
                            streamWriter.WriteLine("Warning: Unable to import attachments for release " + releaseId + " (" + exception.Message + ")");
                        }
                    }

                    //Add to the mapping hashtables
                    if (!this.releaseMapping.ContainsKey(releaseId))
                    {
                        this.releaseMapping.Add(releaseId, newReleaseId);
                    }

                    //Add any iterations (if there are any)
                    if (releaseObject.CycleFactory != null)
                    {
                        ImportIterations(streamWriter, releaseObject.CycleFactory, newReleaseId);
                    }
                }
            }
        }

        /// <summary>
        /// Imports releases into Spira
        /// </summary>
        /// <param name="streamWriter">The logfile writer</param>
        /// <param name="releaseFactory">The QC release factory object</param>
        /// <param name="parentReleaseId">The parent Spira release, or null for top-level</param>
        protected void ImportIterations(StreamWriter streamWriter, HP.ALM.CycleFactory cycleFactory, int? parentReleaseId)
        {
            HP.ALM.TDFilter tdFilter = (HP.ALM.TDFilter)cycleFactory.Filter;
            tdFilter.set_Order("rcyc_name", 1);
            HP.ALM.List iterationList = (HP.ALM.List)tdFilter.NewList();

            //Import the iterations into Spira
            foreach (HP.ALM.Cycle cycleObject in iterationList)
            {
                //Extract the iteration info
                int cycleId = (int)cycleObject.ID;

                //The factory object returns all children, not just the direct ones, so make sure we only load an iteration once
                if (!iterationMapping.ContainsKey(cycleId))
                {
                    //Load the iteration and capture the ID
                    RemoteRelease remoteRelease = new RemoteRelease();
                    remoteRelease.Name = cycleObject.Name;
                    if (cycleObject["rcyc_description"] != null)
                    {
                        remoteRelease.Description = MakeXmlSafe(cycleObject["rcyc_description"]);
                    }
                    remoteRelease.VersionNumber = cycleObject.Name.SafeSubstring(10);
                    remoteRelease.StartDate = cycleObject.StartDate.ToUniversalTime();
                    remoteRelease.EndDate = cycleObject.EndDate.ToUniversalTime();
                    remoteRelease.ResourceCount = 1;
                    remoteRelease.Active = true;
                    remoteRelease.Iteration = true;
                    int newIterationId = ImportFormHandle.SpiraImportProxy.Release_Create(remoteRelease, parentReleaseId).ReleaseId.Value;
                    streamWriter.WriteLine("Added iteration(cycle): '" + cycleId);

                    //Add attachments if requested
                    if (this.MainFormHandle.chkImportAttachments.Checked)
                    {
                        try
                        {
                            ImportAttachments(streamWriter, "RELEASES", cycleId, releaseObject.Attachments, newReleaseId, Utils.ArtifactType.Release);
                        }
                        catch (Exception exception)
                        {
                            streamWriter.WriteLine("Warning: Unable to import attachments for release " + releaseId + " (" + exception.Message + ")");
                        }
                    }

                    //Add to the mapping hashtables
                    if (!this.iterationMapping.ContainsKey(cycleId))
                    {
                        this.iterationMapping.Add(cycleId, newIterationId);
                    }
                }
            }
        }*/

        /*
        /// <summary>
        /// Recursively imports a list of subjects into SpiraTest as test folders
        /// </summary>
        /// <param name="treeManager">Handle to the tree manager</param>
        /// <param name="subjectList">The subject list</param>
        /// <param name="parentFolderId">The SpiraTest parent test folder</param>
        protected void ImportSubjectNode(TreeManager treeManager, HP.ALM.IList subjectList, Nullable<int> parentFolderId)
        {
            for (int i = 1; i <= subjectList.Count; i++)
            {
                //See if we have a subject node or a subject name
                HP.ALM.SubjectNode subjectNode;
                if (subjectList[i].GetType() == typeof(string))
                {
                    //Extract the test folder info from the subject tree
                    string subjectName = (string)subjectList[i];
                    subjectNode = (HP.ALM.SubjectNode)treeManager[subjectName];
                }
                else
                {
                    subjectNode = (HP.ALM.SubjectNode)subjectList[i];
                }

                //Extract the subject node information
                int subjectId = subjectNode.NodeID;
                string name = MakeXmlSafe(subjectNode.Name);

                //Load the test folder and capture the new id
                RemoteTestCase remoteTestFolder = new RemoteTestCase();
                remoteTestFolder.Name = name;
                int newTestFolderId = ImportFormHandle.SpiraImportProxy.TestCase_CreateFolder(remoteTestFolder, parentFolderId).TestCaseId.Value;

                //Add to the mapping hashtable
                this.testFolderMapping.Add(subjectId, newTestFolderId);

                //Now get the children and load them under this node
                HP.ALM.IList childList = subjectNode.NewList();
                if (childList.Count > 0)
                {
                    ImportSubjectNode(treeManager, childList, newTestFolderId);
                }
            }
        }*/

        /*
        /// <summary>
        /// Recursively imports a list of test set folders into SpiraTest as test set folders
        /// </summary>
        /// <param name="treeManager">Handle to the tree manager</param>
        /// <param name="folderList">The subject list</param>
        /// <param name="parentFolderId">The SpiraTest parent test set folder</param>
        protected void ImportTestSetFolderNode(HP.ALM.TestSetTreeManager treeManager, HP.ALM.IList folderList, Nullable<int> parentFolderId)
        {
            for (int i = 1; i <= folderList.Count; i++)
            {
                //See if we have a test set foldernode or a test set folder name
                HP.ALM.TestSetFolder testSetFolderNode;
                testSetFolderNode = (HP.ALM.TestSetFolder)folderList[i];

                //Extract the subject node information
                int testSetFolderNodeId = testSetFolderNode.NodeID;
                string name = MakeXmlSafe(testSetFolderNode.Name);

                //Load the test set folder and capture the new id
                RemoteTestSet remoteTestSetFolder = new RemoteTestSet();
                remoteTestSetFolder.Name = name;
                int newTestSetFolderId = ImportFormHandle.SpiraImportProxy.TestSet_CreateFolder(remoteTestSetFolder, parentFolderId).TestSetId.Value;

                //Add to the mapping hashtable
                this.testSetFolderMapping.Add(testSetFolderNodeId, newTestSetFolderId);

                //Now get the children and load them under this node
                HP.ALM.IList childList = testSetFolderNode.NewList();
                if (childList.Count > 0)
                {
                    ImportTestSetFolderNode(treeManager, childList, newTestSetFolderId);
                }
            }
        }*/

        private void ProgressForm_Load(object sender, EventArgs e)
        {

        }
    }
}
