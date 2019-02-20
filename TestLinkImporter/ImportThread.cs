using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;

using Inflectra.SpiraTest.AddOns.TestLinkImporter.SpiraSoapService;
using Meyn.TestLink;

namespace Inflectra.SpiraTest.AddOns.TestLinkImporter
{
    public class ImportThread
    {
        //Project role for new users
        private const int PROJECT_ROLE_ID = 5;

        protected ProgressForm progressForm;
        protected int testLinkProjectId;
        protected string testLinkProjectName;

        protected Dictionary<int, int> testSuiteMapping;
        protected Dictionary<int, int> testSuiteSectionMapping;
        protected Dictionary<int, int> testCaseMapping;
        protected Dictionary<int, int> testStepMapping;
        protected Dictionary<int, int> testSetMapping;
        protected Dictionary<int, int> testSetTestCaseMapping;
        protected Dictionary<int, int> usersMapping;
        protected Dictionary<int, int> testRunStepMapping;
        protected Dictionary<int, int> testRunStepToTestStepMapping;
        protected Dictionary<int, int> releaseMapping = new Dictionary<int, int>();
        protected int newProjectId;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="progressForm">Handle to the parent form</param>
        public ImportThread(ProgressForm progressForm, int testRailProjectId, string testLinkProjectName)
        {
            this.progressForm = progressForm;
            this.testLinkProjectId = testRailProjectId;
            this.testLinkProjectName = testLinkProjectName;
        }

        /// <summary>
        /// Configure the SOAP connection for HTTP or HTTPS depending on what was specified
        /// </summary>
        /// <param name="httpBinding"></param>
        /// <param name="uri"></param>
        /// <remarks>Allows self-signed certs to be used</remarks>
        public void ConfigureBinding(BasicHttpBinding httpBinding, Uri uri)
        {
            //Handle SSL if necessary
            if (uri.Scheme == "https")
            {
                httpBinding.Security.Mode = BasicHttpSecurityMode.Transport;
                httpBinding.Security.Transport.ClientCredentialType = HttpClientCredentialType.None;

                //Allow self-signed certificates
                PermissiveCertificatePolicy.Enact("");
            }
            else
            {
                httpBinding.Security.Mode = BasicHttpSecurityMode.None;
            }
            httpBinding.AllowCookies = true;
        }

        private void ImportProject(StreamWriter streamWriter, SpiraSoapService.SoapServiceClient spiraClient, TestLink testLinkApi)
        {
            try
            {
                streamWriter.WriteLine(String.Format("Importing Test Link project '{0}' to Spira...", this.testLinkProjectName));

                //Get the project info from the TestLink API
                TestProject tlProject = testLinkApi.GetProject(this.testLinkProjectName);
                if (tlProject == null)
                {
                    throw new ApplicationException(String.Format("Empty project data returned by TestLink for project '{0}' so aborting import", this.testLinkProjectId));
                }

                RemoteProject remoteProject = new RemoteProject();
                remoteProject.Name = tlProject.name;
                remoteProject.Website = tlProject.notes;
                remoteProject.Active = tlProject.active;

                //Reconnect and import the project
                spiraClient.Connection_Authenticate2(Properties.Settings.Default.SpiraUserName, Properties.Settings.Default.SpiraPassword, "TestLinkImporter");
                remoteProject = spiraClient.Project_Create(remoteProject, null);
                streamWriter.WriteLine("New Project '" + remoteProject.Name + "' Created");
                int projectId = remoteProject.ProjectId.Value;

                this.newProjectId = remoteProject.ProjectId.Value;
            }
            catch (Exception exception)
            {
                streamWriter.WriteLine("General error importing data from TestLink to Spira The error message is: " + exception.Message);
                throw;
            }
        }

        /// <summary>
        /// Converts Test Rail execution statuses
        /// </summary>
        private int ConvertExecutionStatus(int testRailStatusId)
        {
            int executionStatusId;
            switch (testRailStatusId)
            {
                case 1:
                    executionStatusId = (int)ExecutionStatusEnum.Passed;
                    break;
                case 2:
                    executionStatusId = (int)ExecutionStatusEnum.Blocked;
                    break;
                case 3:
                    /* Untested*/
                    executionStatusId = (int)ExecutionStatusEnum.NotRun;
                    break;
                case 4:
                    /* Retest*/
                    executionStatusId = (int)ExecutionStatusEnum.Caution;
                    break;
                case 5:
                    /* Failed*/
                    executionStatusId = (int)ExecutionStatusEnum.Failed;
                    break;

                default:
                    //Custom statuses become N/A
                    executionStatusId = (int)ExecutionStatusEnum.NotApplicable;
                    break;

            }

            return executionStatusId;
        }

        /// <summary>
        /// The execution statuses
        /// </summary>
        public enum ExecutionStatusEnum
        {
            Failed = 1,
            Passed = 2,
            NotRun = 3,
            NotApplicable = 4,
            Blocked = 5,
            Caution = 6
        }

        /// <summary>
        /// Imports the test suites and sections
        /// </summary>
        /// <param name="streamWriter"></param>
        /// <param name="spiraClient"></param>
        /// <param name="testRailApi"></param>
        private void ImportTestSuites(StreamWriter streamWriter, SpiraSoapService.SoapServiceClient spiraClient, TestLink testLinkApi, List<TestSuite> testSuites, int? parentTestFolderId = null)
        {
            //Get the test suites from the TestLink API
            if (testSuites != null)
            {
                foreach (TestSuite testSuite in testSuites)
                {
                    //Extract the user data
                    int testLinkId = testSuite.id;

                    //Create the new SpiraTest test folder
                    RemoteTestCaseFolder remoteTestFolder = new RemoteTestCaseFolder();
                    remoteTestFolder.Name = testSuite.name;
                    remoteTestFolder.ParentTestCaseFolderId = parentTestFolderId;
                    int newTestCaseFolderId = spiraClient.TestCase_CreateFolder(remoteTestFolder).TestCaseFolderId.Value;
                    streamWriter.WriteLine("Added test suite: " + testLinkId);

                    //Add to the mapping hashtables
                    if (!this.testSuiteMapping.ContainsKey(testLinkId))
                    {
                        this.testSuiteMapping.Add(testLinkId, newTestCaseFolderId);
                    }

                    //See if this test suite has any child test suites
                    List<TestSuite> childTestSuites = testLinkApi.GetTestSuitesForTestSuite(testLinkId);
                    if (childTestSuites != null && childTestSuites.Count > 0)
                    {
                        ImportTestSuites(streamWriter, spiraClient, testLinkApi, childTestSuites, newTestCaseFolderId);
                    }
                }
            }
        }

        /// <summary>
        /// Imports test cases and associated steps
        /// </summary>
        private void ImportTestCasesAndSteps(StreamWriter streamWriter, SpiraSoapService.SoapServiceClient spiraClient, TestLink testLinkApi)
        {
            //Loop through all the mapped, imported test suites
            foreach (KeyValuePair<int, int> kvp in this.testSuiteMapping)
            {
                //Get the test cases from the TestLink API
                int testSuiteId = kvp.Key;
                int testFolderId = kvp.Value;
                streamWriter.WriteLine(String.Format("Getting test cases in test suite: {0}", testSuiteId));
                List<TestCaseFromTestSuite> testCasesInSuite = testLinkApi.GetTestCasesForTestSuite(testSuiteId, true);
                if (testCasesInSuite != null)
                {
                    foreach (TestCaseFromTestSuite testCaseInSuite in testCasesInSuite)
                    {
                        //Extract the user data
                        int tlTestCaseId = testCaseInSuite.id;
                        string tlTestCaseExternalId = testCaseInSuite.external_id;
                        streamWriter.WriteLine(String.Format("Finding test case: {0} - {1}", tlTestCaseId, tlTestCaseExternalId));

                        //Get the actual test case object (with steps)
                        try
                        {
                            TestCase testCase = testLinkApi.GetTestCase(tlTestCaseId);
                            if (testCase != null)
                            {
                                streamWriter.WriteLine(String.Format("Importing test case: {0}", tlTestCaseId));
                                string description = "";
                                if (!String.IsNullOrEmpty(testCase.summary))
                                {
                                    description += "<p>" + testCase.summary + "</p>";
                                }
                                if (!String.IsNullOrEmpty(testCase.preconditions))
                                {
                                    description += "<p>" + testCase.preconditions + "</p>";
                                }

                                //Reauthenticate at this point, in case disconnected
                                spiraClient.Connection_Authenticate2(Properties.Settings.Default.SpiraUserName, Properties.Settings.Default.SpiraPassword, "TestLinkImporter");
                                spiraClient.Connection_ConnectToProject(this.newProjectId);

                                //Create the new SpiraTest test case
                                RemoteTestCase remoteTestCase = new RemoteTestCase();
                                remoteTestCase.Name = testCase.name;
                                remoteTestCase.Description = description;
                                remoteTestCase.TestCaseTypeId = /* Functional */ 3;
                                remoteTestCase.TestCaseStatusId = (testCase.active) ? /* Ready For Test */ 5 : /* Draft*/ 1;
                                remoteTestCase.CreationDate = testCase.creation_ts;
                                if (testCase.importance >= 1 && testCase.importance <= 4)
                                {
                                    remoteTestCase.TestCasePriorityId = testCase.importance;
                                }

                                //We don't import the author currently
                                //remoteTestCase.AuthorId = this.usersMapping[createdById.Value];

                                //Map to the folder
                                remoteTestCase.TestCaseFolderId = testFolderId;

                                int newTestCaseId = spiraClient.TestCase_Create(remoteTestCase).TestCaseId.Value;
                                streamWriter.WriteLine("Added test case: " + tlTestCaseId);

                                //See if we have any test steps for this test case
                                if (testCase.steps != null && testCase.steps.Count > 0)
                                {
                                    streamWriter.WriteLine("Adding " + testCase.steps.Count + " test steps to test case: " + tlTestCaseId);
                                    foreach (TestStep testStep in testCase.steps)
                                    {
                                        RemoteTestStep remoteTestStep = new RemoteTestStep();
                                        remoteTestStep.Description = testStep.actions;
                                        remoteTestStep.ExpectedResult = testStep.expected_results;
                                        spiraClient.TestCase_AddStep(remoteTestStep, newTestCaseId);
                                    }
                                    streamWriter.WriteLine("Added test steps to test case: " + tlTestCaseId);
                                }

                                //Add to the mapping hashtables
                                if (!this.testCaseMapping.ContainsKey(tlTestCaseId))
                                {
                                    this.testCaseMapping.Add(tlTestCaseId, newTestCaseId);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            streamWriter.WriteLine(String.Format ("Error importing test case {0}: {1} ({2})", tlTestCaseId, ex.Message, ex.StackTrace));
                        }
                    }
                }
            }
        }

        /// <summary>
        /// This method is responsible for actually importing the data
        /// </summary>
        /// <param name="stateInfo">State information handle</param>
        /// <remarks>This runs in background thread to avoid freezing the progress form</remarks>
        public void ImportData(object stateInfo)
        {
            //First open up the textfile that we will log information to (used for debugging purposes)
            string debugFile = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\Spira_TestLink_Import.log";
            StreamWriter streamWriter = File.CreateText(debugFile);

            try
            {
                streamWriter.WriteLine("Starting import at: " + DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString());

                //Create the mapping hashtables
                this.testCaseMapping = new Dictionary<int, int>();
                this.testSuiteMapping = new Dictionary<int, int>();
                this.testSuiteSectionMapping = new Dictionary<int, int>();
                this.testStepMapping = new Dictionary<int, int>();
                this.usersMapping = new Dictionary<int, int>();
                this.testRunStepMapping = new Dictionary<int, int>();
                this.testRunStepToTestStepMapping = new Dictionary<int, int>();
                this.testSetMapping = new Dictionary<int, int>();
                this.testSetTestCaseMapping = new Dictionary<int, int>();

                //Connect to Spira
                streamWriter.WriteLine("Connecting to Spira...");
                SpiraSoapService.SoapServiceClient spiraClient = new SpiraSoapService.SoapServiceClient();

                //Set the end-point and allow cookies
                string spiraUrl = Properties.Settings.Default.SpiraUrl + ImportForm.IMPORT_WEB_SERVICES_URL;
                spiraClient.Endpoint.Address = new EndpointAddress(spiraUrl);
                BasicHttpBinding httpBinding = (BasicHttpBinding)spiraClient.Endpoint.Binding;
                ConfigureBinding(httpBinding, spiraClient.Endpoint.Address.Uri);

                bool success = spiraClient.Connection_Authenticate2(Properties.Settings.Default.SpiraUserName, Properties.Settings.Default.SpiraPassword, "TestLinkImporter");
                if (!success)
                {
                    string message = "Failed to authenticate with Spira using login: '" + Properties.Settings.Default.SpiraUserName + "' so terminating import!";
                    streamWriter.WriteLine(message);
                    streamWriter.Close();

                    //Display the exception message
                    this.progressForm.ProgressForm_OnError(message);
                    return;
                }

                //Connect to TestLink
                string apiKey = Properties.Settings.Default.TestLinkApiKey;
                string testLinkUrl = Properties.Settings.Default.TestLinkUrl + Utils.TESTLINK_URL_SUFFIX_XML_RPC;
                TestLink testLink = new TestLink(apiKey, testLinkUrl);

                //1) Create a new project
                ImportProject(streamWriter, spiraClient, testLink);

                //**** Show progress ****
                streamWriter.WriteLine("Project Created");
                this.progressForm.ProgressForm_OnProgressUpdate(1);

                //2) Import test suites
                //Get the test suites from the TestLink API
                List<TestSuite> testSuites = testLink.GetFirstLevelTestSuitesForTestProject(this.testLinkProjectId);
                ImportTestSuites(streamWriter, spiraClient, testLink, testSuites);

                //**** Show progress ****
                streamWriter.WriteLine("Imported Test Suites");
                this.progressForm.ProgressForm_OnProgressUpdate(2);

                //**** Show progress ****
                streamWriter.WriteLine("Imported Test Cases and Steps");
                this.progressForm.ProgressForm_OnProgressUpdate(3);

                //3) Import test cases and test steps
                ImportTestCasesAndSteps(streamWriter, spiraClient, testLink);

                //**** Show progress ****
                streamWriter.WriteLine("Completed Import");
                this.progressForm.ProgressForm_OnProgressUpdate(4);

                //**** Mark the form as finished ****
                streamWriter.WriteLine("Import completed at: " + DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString());
                this.progressForm.ProgressForm_OnFinish();

                //Close the debugging file
                streamWriter.Close();
            }
            catch (Exception exception)
            {
                //Log the error
                streamWriter.WriteLine("*ERROR* Occurred during Import: '" + exception.Message + "' at " + exception.Source + " (" + exception.StackTrace + ")");
                streamWriter.Close();

                //Display the exception message
                this.progressForm.ProgressForm_OnError(exception);
            }
        }
    }
}
