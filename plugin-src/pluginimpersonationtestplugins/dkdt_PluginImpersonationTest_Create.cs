using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace PluginImpersonationTest.Plugins
{
    /// <summary>
    /// Plugin development guide: https://docs.microsoft.com/powerapps/developer/common-data-impersonatedOrganizationService/plug-ins
    /// Best practices and guidance: https://docs.microsoft.com/powerapps/developer/common-data-impersonatedOrganizationService/best-practices/business-logic/
    /// </summary>
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Style", "IDE1006:Naming Styles", Justification = "Readability due to Dataverse naming")]
    public class dkdt_PluginImpersonationTest_Create : PluginBase
    {
        private const string FIRST_NAME = "firstname";
        private const string LAST_NAME = "lastname";

        public dkdt_PluginImpersonationTest_Create(string unsecureConfiguration, string secureConfiguration)
            : base(typeof(dkdt_PluginImpersonationTest_Create))
        {
        }

        protected override void ExecuteDataversePlugin(ILocalPluginContext localPluginContext)
        {
            if (localPluginContext == null)
            {
                throw new ArgumentNullException(nameof(localPluginContext));
            }

            // Get the organization service to retrieve the environment variable
            var userOrganizationService = localPluginContext.PluginUserService;
            var systemUserIdToImpersonate = GetSystemUserIdToImpersonate(userOrganizationService);

            // Create a new impersonatedOrganizationService to impersonate the user
            var serviceFactory = (IOrganizationServiceFactory)localPluginContext.ServiceProvider.GetService(typeof(IOrganizationServiceFactory));
            var impersonatedOrganizationService = serviceFactory.CreateOrganizationService(systemUserIdToImpersonate);

            // make your calls with impersonatedOrganizationService
            // In the trivial (and yes contrived for simplicity) example below, imagine that the calling user doesn't have access ot the systemuser table, but the impersonated user does
            var columns = new ColumnSet(new string[] { FIRST_NAME, LAST_NAME });
            var systemUser = impersonatedOrganizationService.Retrieve("systemuser", systemUserIdToImpersonate, columns);

            var context = localPluginContext.PluginExecutionContext;
            var entity = (Entity)context.InputParameters["Target"];
            entity["dkdt_systemuserimpersonated"] = $"{systemUser[FIRST_NAME]} {systemUser[LAST_NAME]} - {systemUserIdToImpersonate.ToString()}";
        }

        private static Guid GetSystemUserIdToImpersonate(IOrganizationService organizationService)
        {
            var environmentVariableQuery = new QueryExpression("environmentvariabledefinition")
            {
                ColumnSet = new ColumnSet("environmentvariabledefinitionid"),
                LinkEntities =
                        {
                            new LinkEntity
                            {
                                JoinOperator = JoinOperator.LeftOuter,
                                LinkFromEntityName = "environmentvariabledefinition",
                                LinkFromAttributeName = "environmentvariabledefinitionid",
                                LinkToEntityName = "environmentvariablevalue",
                                LinkToAttributeName = "environmentvariabledefinitionid",
                                Columns = new ColumnSet("value"),
                                EntityAlias = "v"
                            }
                        },
                Criteria = new FilterExpression()
                {
                    Conditions =
                    {
                        new ConditionExpression("schemaname", ConditionOperator.Equal, "dkdt_pluginimpersonationtest_SystemUserIdToImpersonate")
                    }
                }
            };


            var environmentVariables = organizationService.RetrieveMultiple(environmentVariableQuery);
            var systemUserId = Guid.Parse(environmentVariables.Entities[0].GetAttributeValue<AliasedValue>("v.value").Value.ToString());

            return systemUserId;
        }
    }
}
