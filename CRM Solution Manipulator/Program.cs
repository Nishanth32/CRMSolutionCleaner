using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.ServiceModel.Description;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Client.Services;
using Microsoft.Xrm.Client;

namespace CRM_Solution_Manipulator
{
    class Program
    {

        const string EntitySolutionLogicalName = "solution";
         const string EntitySolnCompLogicalName = "solutioncomponent";
        private static OrganizationServiceProxy _serviceProxy = null;
        public static IOrganizationService _service { get; private set; }

        public static IOrganizationService CreateConnection()
        {
            ClientCredentials clinet = new ClientCredentials();
            clinet.UserName.UserName = ConfigurationManager.AppSettings.Get("UserName");
            clinet.UserName.Password = ConfigurationManager.AppSettings.Get("Password");
            string orgUrl = ConfigurationManager.ConnectionStrings["CRMConnection"].ConnectionString;
            Uri uri = new Uri(orgUrl);
            using (_serviceProxy = new OrganizationServiceProxy(uri, null, clinet, null))
            {
                _serviceProxy.Timeout = TimeSpan.FromDays(2);
                _service = (IOrganizationService)_serviceProxy;
            }
            Console.WriteLine("Connection to CRM Online Created.");
            return _service;
        }


        public static EntityCollection RetrieveAllSolutions()
        {
            QueryExpression solutionQuery = new QueryExpression
            {
                EntityName = EntitySolutionLogicalName,
                ColumnSet = new ColumnSet(new string[] { "ismanaged", "uniquename", "version", "friendlyname","installedon"}),                
                Criteria = new FilterExpression()
            };
            solutionQuery.Criteria.AddCondition("ismanaged", ConditionOperator.Equal,false);
            return _service.RetrieveMultiple(solutionQuery);
        }

        public static string GetSolutionNameFromUser(List<string> availableEntities)
        {
               while(true)
            {
                Console.WriteLine("\n Please Enter the name of the Solution to empty : ");
                string userInput = Convert.ToString(Console.ReadLine());
                var datFound = availableEntities.Where(x => x.ToLower() == userInput.ToLower()).FirstOrDefault();
                if(datFound!=null)
                {
                    return datFound;
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Solution not found. Please try again"); Console.ForegroundColor = ConsoleColor.White;
                }
            }
        }


        public static EntityCollection GrabAllSoluionComponent(Guid solutionID)
        {
            QueryByAttribute componentQuery = new QueryByAttribute
            {
                EntityName = EntitySolnCompLogicalName,
                ColumnSet = new ColumnSet("componenttype", "objectid", "solutioncomponentid", "solutionid"),               
                Attributes = { "solutionid" },               
                Values = { solutionID }
            };

            return _service.RetrieveMultiple(componentQuery);
        }

        public static void RemoveAllSolutionComponent(EntityCollection solutionComponents, string solnUniqueName)
        {
            try
            {
                if (solutionComponents == null || solutionComponents.Entities.Count == 0) return;
                Console.ForegroundColor = ConsoleColor.Green; Console.WriteLine("Preparing to empty Solution : " + solnUniqueName);Console.ForegroundColor = ConsoleColor.White;

                foreach(var component in solutionComponents.Entities)
                {
                    Console.ForegroundColor = ConsoleColor.Blue;
                    try
                    {                                                
                        RemoveSolutionComponentRequest removeReq = new RemoveSolutionComponentRequest()
                        {
                            ComponentId =component.GetAttributeValue<Guid>("objectid"),
                            ComponentType = component.GetAttributeValue<OptionSetValue>("componenttype").Value,
                            SolutionUniqueName = solnUniqueName
                        };
                        _serviceProxy.Execute(removeReq);
                    }
                    catch (Exception ex)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine(Environment.NewLine + "\nFailed to Remove Solution Component - {0}", component.GetAttributeValue<Guid>("objectid").ToString().PadRight(20, '.'));
                        Console.ForegroundColor = ConsoleColor.White;
                    }
                    Console.WriteLine(Environment.NewLine + "Removed Solution Component - {0}", component.GetAttributeValue<Guid>("objectid").ToString().PadRight(20, '.'));
                }
                Console.ForegroundColor = ConsoleColor.Green; Console.WriteLine("Completed {0} truncation with above status", solnUniqueName); Console.ForegroundColor = ConsoleColor.White;
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Access to remove solution component failed");
                Console.ForegroundColor = ConsoleColor.White;
            }
            finally
            {
                Console.ForegroundColor = ConsoleColor.White;
            }
        }

        public static void DisplayAllSolutions(EntityCollection collection)
        {
            if (collection == null || collection.Entities.Count == 0) return;
            Console.WriteLine("\r\n List of UnManaged Solutions Found:");
            int i = 0;
            Console.WriteLine("{0}. {1} {2} {3}", "S.No", "Solution Name".PadRight(40),"Version".PadRight(20),"InstalledOn".PadRight(10));
            foreach (var entity in collection.Entities)
            {
                Console.WriteLine("{0} {1} {2} {3}", ((++i).ToString()).PadRight(5),entity.GetAttributeValue<string>("uniquename").PadRight(40), entity.GetAttributeValue<string>("version").PadRight(20), entity.GetAttributeValue<DateTime>("installedon").ToString().PadRight(10));
            }
        }

        static void Main(string[] args)
        {

            string vb = "dev1_/RWB2013/RibbonWorkbench.xap";

            QueryExpression qe = new QueryExpression("entityName");
            qe.ColumnSet = new ColumnSet(true);
            FilterExpression filter1 = new FilterExpression(LogicalOperator.And);
            filter1.Conditions.Add(new ConditionExpression("A_LogicalName", ConditionOperator.Equal, "100002"));
            qe.Criteria = filter1;


            string output = new System.Web.Script.Serialization.JavaScriptSerializer().Serialize(qe);
            Console.WriteLine(output);
            try
            {
                //CreateConnection();
                CreateTempConnection();
                while (true)
                {
                  Console.Write("Retrieving Solutions..."); Console.Write("\r");
                 var solutionCollection = RetrieveAllSolutions();
                    Console.WriteLine("".PadRight(100, '.'));
                    DisplayAllSolutions(solutionCollection);
                    string solutionUniqueName = GetSolutionNameFromUser(solutionCollection.Entities.Select(x => x.GetAttributeValue<string>("uniquename")).ToList());

                    Guid targetSolutionGuid = solutionCollection.Entities.Where(x => x.GetAttributeValue<string>("uniquename") == solutionUniqueName).Select(y => y.GetAttributeValue<Guid>("solutionid")).First();
                    // Get all the Solution Components
                    var solutionComponents = GrabAllSoluionComponent(targetSolutionGuid);
                    if(solutionComponents.Entities.Count == 0)
                    {
                        Console.WriteLine("Selected solution is already empty.");
                        continue;
                    }
                    RemoveAllSolutionComponent(solutionComponents,solutionUniqueName);
                    Console.WriteLine("".PadRight(100, '.'));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Code Failed to Execute" + ex.Message);
            }          
        }


        static void CreateTempConnection()
        {
            string connectionString = "Url=http://vsevm.centralus.cloudapp.azure.com/xypex; Domain=JLDOMAIN; Username=Antony.Nishanth; Password=U^1$973vmd3D1>H;";
            CrmConnection connection = CrmConnection.Parse(connectionString);

            _service = new OrganizationService(connection); 
        }
    }
}
