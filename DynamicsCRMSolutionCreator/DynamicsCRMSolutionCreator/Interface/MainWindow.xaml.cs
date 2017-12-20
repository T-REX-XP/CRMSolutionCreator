using Microsoft.Xrm.Client.Services;
using System;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Xrm.Client;
using System.ServiceModel;
using System.IO;
using System.Data;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Discovery;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.VisualBasic;
using GenericParsing;
using DynamicsCRMSolutionCreator;

namespace DynamicsCRMSolutionCreator
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        public void AppendTextBox(string textvalue)
        {
            string linefeed="\n";
            Dispatcher.Invoke(new Action(() => txtboxStatusDetails.AppendText(textvalue + linefeed)), DispatcherPriority.Render);
            txtboxStatusDetails.ScrollToEnd();
        }
        public void SetStatus(string statusvalue, SolidColorBrush textcolor)
        {
            Dispatcher.Invoke(new Action(() => txtStatus.Content=(string)statusvalue), DispatcherPriority.Render);
            Dispatcher.Invoke(new Action(() => txtStatus.Foreground=textcolor), DispatcherPriority.Render);
        }
        private void btnCreateSolution_click(object sender, RoutedEventArgs e)
        {
            ////Clear the details text box
            txtboxStatusDetails.Clear();
            txtStatus.Content = (string)"Processing";
            txtStatus.Foreground=Brushes.Black;
            
            ////Retrieve the text from the CSV location and check if it exists
            string path = (string)txtCSVLocation.Text;
            List<string[]> componentlist;
            if (!File.Exists(path))
            {
                AppendTextBox("Error: File path does not exist");
                SetStatus("ERROR", Brushes.Red);
                return;
            }
            else
            {
                    AppendTextBox("Extracting CSV data");
                    componentlist=(List<string[]>)ReadCSVFile(path);
                    AppendTextBox("CSV data obtained");
            }

            ////Retrieve the text from the Solution Name and check if it exists
            string solutionName = txtsolutionName.Text;
            if (!(solutionName.Length > 0))
            {
                AppendTextBox("Error: Solution name was not entered");
                SetStatus("ERROR", Brushes.Red);
                return;
            }

            ////Attempt to connect to CRM
            string crmURL = (string)txtCRMURL.Text;
            string password = (string)txtPassword.Password;
            string username = (string)txtUserName.Text;
            OrganizationService service = ConnectToCrm(username, password, crmURL);
            if (service == null)
            {
                return;
            }

            ////Create solution if does not already exist
            Solution solution=CreateSolution(service,solutionName);
            if (solution == null)
            {
                AppendTextBox("ERROR: Solution create or update failed");
                SetStatus("ERROR", Brushes.Red);
                return;
            }
            AppendTextBox("Solution retrieved");

            ////Add Components to solution
            BeginAddSolutionComponents(service, solutionName, componentlist);
            SetStatus("COMPLETE", Brushes.Blue);
            }
        public OrganizationService ConnectToCrm(string username, string password, string URL)
        {
            try
            {
                CrmConnection connection = new CrmConnection();
                string endserviceLink = (string)"/XRMServices/2011/Organization.svc";
                string fullservicelink = URL + endserviceLink;
                connection.ServiceUri = new Uri(fullservicelink);
                connection.ClientCredentials = new System.ServiceModel.Description.ClientCredentials();
                connection.ClientCredentials.UserName.UserName = username;
                connection.ClientCredentials.UserName.Password = password;
                connection.Timeout = new TimeSpan(0, 10, 0);
                return new OrganizationService(connection);
            }
            catch
            {
                AppendTextBox("ERROR: Could not connect to CRM");
                SetStatus("ERROR", Brushes.Red);
                return new OrganizationService(new CrmConnection());
            }
        }
        public Solution CreateSolution(IOrganizationService service, string solutionName)
        {
            //Define variables
            Solution retrievedSolution = null;
            Guid publisherid = Guid.Empty;

            //Check if solution exists, otherwise create
            AppendTextBox("Retrieving solution: " + solutionName);
            retrievedSolution = QuerySolution(service, solutionName);
            if (retrievedSolution!=null)
            {
                AppendTextBox("Solution exists");
                return retrievedSolution;
            }
            else
            {
                //Query and Define Publisher
                try
                {
                    string publisherName = "H21";
                    QueryExpression queryPublisher = new QueryExpression
                    {
                        EntityName = Publisher.EntityLogicalName,
                        ColumnSet = new ColumnSet("publisherid"),
                        Criteria = new FilterExpression()
                    };
                    queryPublisher.Criteria.AddCondition("uniquename", ConditionOperator.BeginsWith, publisherName);
                    EntityCollection queryPublisherResults = service.RetrieveMultiple(queryPublisher);
                    Publisher retrievedPublisher = (Publisher)queryPublisherResults.Entities[0];
                    publisherid = (Guid)retrievedPublisher.PublisherId;
                }
                catch
                {
                    AppendTextBox("ERROR: Could not retrieve solution publisher that begins with H21");
                    SetStatus("ERROR", Brushes.Red);
                    return retrievedSolution;
                }
                

                //Define a solution
                Solution newsolution = new Solution
                {
                    UniqueName = solutionName,
                    FriendlyName = solutionName,
                    PublisherId = new EntityReference(Publisher.EntityLogicalName, publisherid),
                    Description = "Created using CRM Solution Creator",
                    Version = "1.0.0.0"
                };

                AppendTextBox("Attempting to create solution");

                Guid retrievedSolutionID=service.Create(newsolution);

                retrievedSolution = QuerySolution(service, solutionName);

                return retrievedSolution;
            }
            
        }
        public Solution QuerySolution(IOrganizationService service,string solutionName)
        {
            Solution retrievedSolution = null;
            try
            {
                ////Attempt to retrieve solution
                QueryExpression querySolution = new QueryExpression
                {
                    EntityName = Solution.EntityLogicalName,
                    ColumnSet = new ColumnSet(),
                    Criteria = new FilterExpression()
                };
                querySolution.Criteria.AddCondition("uniquename", ConditionOperator.Equal, solutionName);
                EntityCollection querySolutionResults = service.RetrieveMultiple(querySolution);
                

                ////Check if solution was retrieved
                if (querySolutionResults.Entities.Count > 0)
                {
                    retrievedSolution = (Solution)querySolutionResults.Entities[0];
                    return retrievedSolution;
                }

                ////Return a null solution if not retrieved
                return retrievedSolution;
            }
            catch
            {
                AppendTextBox("Could not retrieve or create solution. Check credentials and CRM URL.");
                SetStatus("ERROR", Brushes.Red);
                return retrievedSolution;
            }
            
        }
        public void BeginAddSolutionComponents(IOrganizationService service, string solutionName,List<string[]> componentlist )
        {
             foreach (string[] component in componentlist)  
             {
                string ObjectName=component[0].ToString();
                string ObjectType=component[1].ToString();

                AppendTextBox("Attempting to retrieve " + ObjectName + " of type " + ObjectType);
                switch (ObjectType)
                {
                    case "Entity":
                        AddEntity(service,ObjectName,solutionName);
                        break;
                    case "Option Set":
                        AddOptionSet(service,ObjectName,solutionName);
                        break;
                    case "Web Resource":
                        AddWebResource(service,ObjectName,solutionName);
                        break;
                    case "Workflow":
                        AddProcess(service,ObjectName,solutionName);
                        break;
                    case "Dialog":
                        AddProcess(service,ObjectName,solutionName);
                        break;
                    case "Action":
                        AddProcess(service,ObjectName,solutionName);
                        break;
                    case "Security Role":
                        AddRole(service,ObjectName,solutionName);
                        break;
                    case "SDK Message Processing Step":
                        AddSDKStep(service,ObjectName,solutionName);
                        break;
                    case "Workflow Assembly":
                        AddAssembly(service,ObjectName,solutionName);
                        break;
                    case "Plugin":
                        AddAssembly(service,ObjectName,solutionName);
                        break;
                    default:
                        AppendTextBox(ObjectType + " is not yet supported");
                        break;
                }
             }
        }
        public List<string[]> ReadCSVFile(string path)
        {
            List<string[]> components = new List<string[]>();
            try
            {
                ///Set variables to store values in
                string ObjectName;
                string ObjectType;
                using (GenericParser parser = new GenericParser())
                {
                    parser.SetDataSource(path);
                    parser.ColumnDelimiter = ',';
                    parser.FirstRowHasHeader = true;
                    parser.MaxBufferSize = 4096;
                    parser.MaxRows = 50000;
                    parser.TextQualifier = '\"';
                    parser.SkipEmptyRows = true;

                    while (parser.Read())
                    {
                        ObjectName = parser["Object Name"];
                        ObjectType = parser["Object Type"];

                        string[] component = new string[] { ObjectName, ObjectType };
                        components.Add(component);

                    }
                    return components;
                }
            }
            catch
            {
                AppendTextBox("Error reading CSV file");
                SetStatus("ERROR", Brushes.Red);
                return components;
            }

        }
        public void AddComponent (IOrganizationService service, string solutionName, int componentType,Guid componentid)
        {
            AppendTextBox("Attempting to add component to " + solutionName);
            AddSolutionComponentRequest addComponentRequest = new AddSolutionComponentRequest()
            {
                ComponentType = componentType,
                ComponentId = componentid,
                SolutionUniqueName = solutionName
            };
            service.Execute(addComponentRequest);
            AppendTextBox("Component added successfully");
            return;
        }

        #region Add Different Component Methods

        private void AddEntity(IOrganizationService service, string ObjectName, string solutionName)
        {
            RetrieveEntityRequest retrieveEntityRequest = new RetrieveEntityRequest()
                {
                    LogicalName = ObjectName
                };
            RetrieveEntityResponse retrieveEntityResponse = (RetrieveEntityResponse)service.Execute(retrieveEntityRequest);

            Guid componentid = (Guid)retrieveEntityResponse.EntityMetadata.MetadataId;
            int componentType = 1;

            AddComponent(service,solutionName,componentType,componentid);

            return;
        }
        private void AddOptionSet(IOrganizationService service, string ObjectName, string solutionName)
        {
            RetrieveOptionSetRequest retrieveOptionSetRequest = new RetrieveOptionSetRequest()
            {
                Name = ObjectName
            };
            RetrieveOptionSetResponse retrieveOptionSetResponse = (RetrieveOptionSetResponse)service.Execute(retrieveOptionSetRequest);

            Guid componentid = (Guid)retrieveOptionSetResponse.OptionSetMetadata.MetadataId;
            int componentType = 9;

            AddComponent(service, solutionName, componentType, componentid);
            return;
        }
        private void AddProcess(IOrganizationService service, string ObjectName, string solutionName)
        {
            ////Define fields to query
            string fieldQuery="name";
            string fieldReturn="workflowid";
            string entity = "workflow";

            //Build query expression
            QueryExpression query = new QueryExpression(entity);
                ColumnSet column = new ColumnSet();
                column.AddColumns(new string[] {fieldReturn});
                query.ColumnSet = column;
                query.Criteria.AddCondition(new ConditionExpression(fieldQuery, ConditionOperator.Equal, ObjectName));
                query.Criteria.AddCondition(new ConditionExpression("type", ConditionOperator.Equal, 1));

            //Retrieve records
            EntityCollection response = service.RetrieveMultiple(query);
            if (response == null || response.Entities.Count == 0)
            {
                AppendTextBox("Error: could not retrieve " + ObjectName);
                return;
            }

            Guid componentid = (Guid)(response.Entities[0].Attributes[fieldReturn]);
            int componentType = 29;

            //Add retrieved component to solution
            AddComponent(service, solutionName, componentType, componentid);
            return;
        }
        private void AddWebResource(IOrganizationService service, string ObjectName, string solutionName)
        {
            ///Define fields to query
            string fieldQuery="name";
            string fieldReturn="webresourceid";
            string entity = "webresource";

            //Build query expression
            QueryExpression query = new QueryExpression(entity);
                ColumnSet column = new ColumnSet();
                column.AddColumns(new string[] {fieldReturn});
                query.ColumnSet = column;
                query.Criteria.AddCondition(new ConditionExpression(fieldQuery, ConditionOperator.Equal, ObjectName));

            //Retrieve records
            EntityCollection response = service.RetrieveMultiple(query);
            if (response == null || response.Entities.Count == 0)
            {
                AppendTextBox("Error: could not retrieve " + ObjectName);
                return;
            }

            Guid componentid = (Guid)(response.Entities[0].Attributes[fieldReturn]);
            int componentType = 61;

            AddComponent(service, solutionName, componentType, componentid);
            return;
        }
        private void AddAssembly(IOrganizationService service, string ObjectName, string solutionName)
        {
            ///Define fields to query
            string fieldQuery = "name";
            string fieldReturn = "pluginassemblyid";
            string entity = "pluginassembly";

            //Build query expression
            QueryExpression query = new QueryExpression(entity);
            ColumnSet column = new ColumnSet();
            column.AddColumns(new string[] { fieldReturn });
            query.ColumnSet = column;
            query.Criteria.AddCondition(new ConditionExpression(fieldQuery, ConditionOperator.Equal, ObjectName));

            //Retrieve records
            EntityCollection response = service.RetrieveMultiple(query);
            if (response == null || response.Entities.Count == 0)
            {
                AppendTextBox("Error: could not retrieve " + ObjectName);
                return;
            }

            Guid componentid = (Guid)(response.Entities[0].Attributes[fieldReturn]);
            int componentType = 91;

            AddComponent(service, solutionName, componentType, componentid);
            return;
        }
        private void AddSDKStep(IOrganizationService service, string ObjectName, string solutionName)
        {
            ///Define fields to query
            string fieldQuery = "name";
            string fieldReturn = "sdkmessageprocessingstepid";
            string entity = "sdkmessageprocessingstep";

            //Build query expression
            QueryExpression query = new QueryExpression(entity);
            ColumnSet column = new ColumnSet();
            column.AddColumns(new string[] { fieldReturn });
            query.ColumnSet = column;
            query.Criteria.AddCondition(new ConditionExpression(fieldQuery, ConditionOperator.Equal, ObjectName));

            //Retrieve records
            EntityCollection response = service.RetrieveMultiple(query);
            if (response == null || response.Entities.Count == 0)
            {
                AppendTextBox("Error: could not retrieve " + ObjectName);
                return;
            }

            Guid componentid = (Guid)(response.Entities[0].Attributes[fieldReturn]);
            int componentType = 92;

            AddComponent(service, solutionName, componentType, componentid);
            return;
        }
        private void AddRole(IOrganizationService service, string ObjectName, string solutionName)
        {
            ///Define fields to query
            string fieldQuery = "name";
            string fieldReturn = "roleid";
            string entity = "role";

            //Build query expression
            QueryExpression query = new QueryExpression(entity);
            ColumnSet column = new ColumnSet();
            column.AddColumns(new string[] { fieldReturn });
            query.ColumnSet = column;
            query.Criteria.AddCondition(new ConditionExpression(fieldQuery, ConditionOperator.Equal, ObjectName));

            //Retrieve records
            EntityCollection response = service.RetrieveMultiple(query);
            if (response == null || response.Entities.Count == 0)
            {
                AppendTextBox("Error: could not retrieve " + ObjectName);
                return;
            }

            Guid componentid = (Guid)(response.Entities[0].Attributes[fieldReturn]);
            int componentType = 20;

            AddComponent(service, solutionName, componentType, componentid);
            return;

        }
        #endregion

        #region Text Changed Clear Details and Status
        
        private void InputChange_TextChanged(object sender, TextChangedEventArgs e)
        {
            ////Replace CSV Location quotes
            TextBox source = (TextBox)sender;
            source.Text=Regex.Replace(source.Text, @"[\""]", "", RegexOptions.None); 

            ////If changed and error then set to WAITING FOR INPUT
            if ((string)txtStatus.Content == "ERROR")
            {
                txtboxStatusDetails.Clear();
                SetStatus("WAITING FOR INPUT", Brushes.Black);
            }
            return;
        }
        private void InputChange_PasswordChanged(object sender, RoutedEventArgs e)
        {
            if ((string)txtStatus.Content == "ERROR")
            {
                txtboxStatusDetails.Clear();
                SetStatus("WAITING FOR INPUT", Brushes.Black);
            }
            return;
        }

        #endregion

        private void ComboBoxItem_Selected(object sender, RoutedEventArgs e)
        {

        }
    }
}



