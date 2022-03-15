using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk.Workflow;
using System.Activities;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace DeleteAccountRelatedRecords
{
    public class DeleteRecords : CodeActivity
    {
        [Input("Property Hold")]
        [RequiredArgument]
        public InArgument<bool> propertyHold { set; get; }
        protected override void Execute(CodeActivityContext executionContext)
        {
            bool allCasesDeleted = false;
            bool allContactsDeleted = false;
            bool allActivitiesDeleted = false;
            bool allAttachmentsDeleted = false;
            bool allEntitlementsDeleted = false;
            int i = 0;

            ITracingService tracingService = executionContext.GetExtension<ITracingService>();

            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory =
                executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service =
                serviceFactory.CreateOrganizationService(context.UserId);
            try
            {
                bool hold = this.propertyHold.Get(executionContext);
                tracingService.Trace("hold value : " + hold);

                Guid accountId = context.PrimaryEntityId;              

                if (hold == true)
                {
                    tracingService.Trace("Records are not deleted as hold is applied");
                    return;
                }

                //Delete all the records related to the account

                //Deleting Cases related to account
                string caseQuery = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                  <entity name='incident'>
                                    <attribute name='title' />
                                    <attribute name='ticketnumber' />
                                    <attribute name='createdon' />
                                    <attribute name='incidentid' />
                                    <attribute name='caseorigincode' />
                                    <order attribute='title' descending='false' />
                                    <link-entity name='account' from='accountid' to='customerid' link-type='inner' alias='ah'>
                                      <filter type='and'>
                                        <condition attribute='accountid' operator='eq' uitype='account' value='{0}' />
                                      </filter>
                                    </link-entity>
                                  </entity>
                                </fetch>";

                caseQuery = string.Format(caseQuery, accountId);
                EntityCollection caseCollection = service.RetrieveMultiple(new FetchExpression(caseQuery));

                if (caseCollection.Entities.Count > 0)
                {
                    foreach (Entity caseRecord in caseCollection.Entities)
                    {
                        service.Delete("incident", caseRecord.Id);
                    }
                    allCasesDeleted = true;
                    tracingService.Trace("All cases have been deleted");
                }
                else
                {
                    allCasesDeleted = true;
                    tracingService.Trace("No case is associated with the account");
                }


                //Deleting Contacts related to account
                string contactQuery = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                      <entity name='contact'>
                                        <attribute name='fullname' />
                                        <attribute name='telephone1' />
                                        <attribute name='contactid' />
                                        <order attribute='fullname' descending='false' />
                                        <link-entity name='account' from='accountid' to='parentcustomerid' link-type='inner' alias='ax'>
                                          <filter type='and'>
                                            <condition attribute='accountid' operator='eq' uitype='account' value='{0}' />
                                          </filter>
                                        </link-entity>
                                      </entity>
                                    </fetch>";

                contactQuery = string.Format(contactQuery, accountId);
                EntityCollection contactCollection = service.RetrieveMultiple(new FetchExpression(contactQuery));

                if (contactCollection.Entities.Count > 0)
                {
                    foreach (Entity contactRecord in contactCollection.Entities)
                    {
                        service.Delete("contact", contactRecord.Id);
                    }
                    allContactsDeleted = true;
                    tracingService.Trace("All contacts have been deleted");
                }
                else
                {
                    allContactsDeleted = true;
                    tracingService.Trace("No contact is associated with the account");
                }
                
                //Deleting Activities related to account
                string activitiesQuery = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                      <entity name='activitypointer'>
                                        <attribute name='activitytypecode' />
                                        <attribute name='subject' />
                                        <attribute name='statecode' />
                                        <attribute name='prioritycode' />
                                        <attribute name='modifiedon' />
                                        <attribute name='activityid' />
                                        <attribute name='instancetypecode' />
                                        <attribute name='community' />
                                        <order attribute='modifiedon' descending='false' />
                                        <link-entity name='account' from='accountid' to='regardingobjectid' link-type='inner' alias='ar'>
                                          <filter type='and'>
                                            <condition attribute='accountid' operator='eq' uitype='account' value='{0}' />
                                          </filter>
                                        </link-entity>
                                      </entity>
                                    </fetch>";

                activitiesQuery = string.Format(activitiesQuery, accountId);
                EntityCollection activitiesCollection = service.RetrieveMultiple(new FetchExpression(activitiesQuery));
                tracingService.Trace("Activities count" + activitiesCollection.Entities.Count);
                if (activitiesCollection.Entities.Count > 0)
                {
                    foreach (Entity activityRecord in activitiesCollection.Entities)
                    {                       
                        string activityTypeCode = activityRecord.GetAttributeValue<string>("activitytypecode").ToLower();                      
                        {
                            switch(activityTypeCode)
                            {
                                case "appointment":
                                    service.Delete("appointment", activityRecord.Id);
                                    tracingService.Trace("Appointment Deleted");
                                    break;

                                case "email":
                                    service.Delete("email", activityRecord.Id);
                                    tracingService.Trace("Email Deleted");
                                    break;

                                case "fax":
                                    service.Delete("fax", activityRecord.Id);
                                    tracingService.Trace("Fax Deleted");
                                    break;

                                case "letter":
                                    service.Delete("letter", activityRecord.Id);
                                    tracingService.Trace("Letter Deleted");
                                    break;

                                case "service activity":
                                    service.Delete("serviceappointment", activityRecord.Id);
                                    tracingService.Trace("Service activity Deleted");
                                    break;

                                case "campaign response":
                                    service.Delete("campaignresponse", activityRecord.Id);
                                    tracingService.Trace("campaign response Deleted");
                                    break;

                                case "phone call":
                                    service.Delete("phonecall", activityRecord.Id);
                                    tracingService.Trace("Phone call Deleted");
                                    break;

                                case "task":
                                    service.Delete("task", activityRecord.Id);
                                    tracingService.Trace("task Deleted");
                                    break;

                                case "booking alert":
                                    service.Delete("msdyn_bookingalert", activityRecord.Id);
                                    tracingService.Trace("Booking Alert Deleted");
                                    break;

                                case "conversation":
                                    service.Delete("msdyn_ocliveworkitem", activityRecord.Id);
                                    tracingService.Trace("Conversation Deleted");
                                    break;

                                case "customer voice alert":
                                    service.Delete("msfp_alert", activityRecord.Id);
                                    tracingService.Trace("Customer Voice alert Deleted");
                                    break;

                                case "outbound message":
                                    service.Delete("msdyn_ocoutboundmessage", activityRecord.Id);
                                    tracingService.Trace("Outbound message Deleted");
                                    break;

                                case "project service approval":
                                    service.Delete("msdyn_approval", activityRecord.Id);
                                    tracingService.Trace("Project service approval Deleted");
                                    break;

                                case "session":
                                    service.Delete("msdyn_ocsession", activityRecord.Id);
                                    tracingService.Trace("Session Deleted");
                                    break;
                            }
                        }
                        
                    }
                    allActivitiesDeleted = true;
                    tracingService.Trace("All activities have been deleted");
                }
                else
                {
                    allActivitiesDeleted = true;
                    tracingService.Trace("No activity is associated with the account");
                } 
                

                //Deleting Notes and Attachments related to account
                string attachmentQuery = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                      <entity name='annotation'>
                                        <attribute name='subject' />
                                        <attribute name='notetext' />
                                        <attribute name='filename' />
                                        <attribute name='annotationid' />
                                        <order attribute='subject' descending='false' />
                                        <link-entity name='account' from='accountid' to='objectid' link-type='inner' alias='at'>
                                          <filter type='and'>
                                            <condition attribute='accountid' operator='eq' uitype='account' value='{0}' />
                                          </filter>
                                        </link-entity>
                                      </entity>
                                    </fetch>";

                attachmentQuery = string.Format(attachmentQuery, accountId);
                EntityCollection attachmentsCollection = service.RetrieveMultiple(new FetchExpression(attachmentQuery));

                if (attachmentsCollection.Entities.Count > 0)
                {
                    foreach (Entity attachmentRecord in attachmentsCollection.Entities)
                    {
                        service.Delete("annotation", attachmentRecord.Id);
                    }
                    allAttachmentsDeleted = true;
                    tracingService.Trace("All notes and attachments have been deleted");
                }
                else
                {
                    allAttachmentsDeleted = true;
                    tracingService.Trace("No attachment is associated with the account");
                }


                //Deleting Entitlements related to account
                string entitlementQuery = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                          <entity name='entitlement'>
                                            <attribute name='name' />
                                            <attribute name='createdon' />
                                            <attribute name='entitytype' />
                                            <attribute name='entitlementid' />
                                            <order attribute='name' descending='false' />
                                            <link-entity name='account' from='accountid' to='customerid' link-type='inner' alias='ai'>
                                              <filter type='and'>
                                                <condition attribute='accountid' operator='eq' uitype='account' value='{0}' />
                                              </filter>
                                            </link-entity>
                                          </entity>
                                        </fetch>";

                entitlementQuery = string.Format(entitlementQuery, accountId);
                EntityCollection entitlementCollection = service.RetrieveMultiple(new FetchExpression(entitlementQuery));

                if (entitlementCollection.Entities.Count > 0)
                {
                    foreach (Entity entitlementRecord in entitlementCollection.Entities)
                    {
                        service.Delete("entitlement", entitlementRecord.Id);
                    }
                    allEntitlementsDeleted = true;
                    tracingService.Trace("All entitlements have been deleted");
                }
                else
                {
                    allEntitlementsDeleted = true;
                    tracingService.Trace("No entitlement is associated with the account");
                }


                // If all related entities to account are deleted then only delete account
                if (allCasesDeleted && allContactsDeleted && allActivitiesDeleted && allAttachmentsDeleted && allEntitlementsDeleted)
                {
                    service.Delete("account", accountId);
                    tracingService.Trace("Account deleted");
                }
            }
            catch (InvalidPluginExecutionException ex)
            {
                throw new InvalidPluginExecutionException(ex.ToString());
            }
        }
    }
}
