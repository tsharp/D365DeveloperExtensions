// =====================================================================
//  This file is based on code from the Microsoft Dynamics CRM SDK code samples.
//
//  Copyright (C) Microsoft Corporation.  All rights reserved.
//
//  This source code is intended only as a supplement to Microsoft
//  Development Tools and/or on-line documentation.  See these other
//  materials for detailed information regarding Microsoft code samples.
//
//  THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY
//  KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
//  IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
//  PARTICULAR PURPOSE.
// =====================================================================

using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Globalization;
using System.ServiceModel;
// ReSharper disable MemberCanBePrivate.Global

namespace $rootnamespace$
{
    /// <inheritdoc />
    /// <summary>
    /// Base class for all workflow activity classes.
    /// </summary> 
    public abstract class WorkFlowActivityBase : CodeActivity
    {
        /// <summary>
        /// Workflow context object. 
        /// </summary>
        protected class LocalWorkflowContext
        {
            internal IServiceProvider ServiceProvider { get; private set; }

            /// <summary>
            /// The Microsoft Dynamics CRM organization service.
            /// </summary>
            internal IOrganizationService OrganizationService { get; }

            /// <summary>
            /// IWorkflowContext contains information that describes the run-time environment in which the workflow executes, information related to the execution pipeline, and entity business information.
            /// </summary>
            internal IWorkflowContext WorkflowExecutionContext { get; }

            /// <summary>
            /// Provides logging run-time trace information for plug-ins. 
            /// </summary>
            internal ITracingService TracingService { get; }

            private LocalWorkflowContext() { }

            /// <summary>
            /// Helper object that stores the services available in this workflow activity.
            /// </summary>
            /// <param name="executionContext"></param>
            internal LocalWorkflowContext(CodeActivityContext executionContext)
            {
                if (executionContext == null)
                {
                    throw new ArgumentNullException(nameof(executionContext));
                }

                // Obtain the execution context service from the service provider.
                WorkflowExecutionContext = executionContext.GetExtension<IWorkflowContext>();

                // Obtain the tracing service from the service provider.
                TracingService = executionContext.GetExtension<ITracingService>();

                // Obtain the Organization Service factory service from the service provider
                IOrganizationServiceFactory factory = executionContext.GetExtension<IOrganizationServiceFactory>();

                // Use the factory to generate the Organization Service.
                OrganizationService = factory.CreateOrganizationService(WorkflowExecutionContext.UserId);
            }

            /// <summary>
            /// Writes a trace message to the CRM trace log.
            /// </summary>
            /// <param name="message">Message name to trace.</param>
            internal void Trace(string message)
            {
                if (string.IsNullOrWhiteSpace(message) || TracingService == null)
                {
                    return;
                }

                if (WorkflowExecutionContext == null)
                {
                    TracingService.Trace(message);
                }
                else
                {
                    TracingService.Trace(
                        "{0}, Correlation Id: {1}, Initiating User: {2}",
                        message,
                        WorkflowExecutionContext.CorrelationId,
                        WorkflowExecutionContext.InitiatingUserId);
                }
            }
        }

        /// <inheritdoc />
        /// <summary>
        /// Main entry point for he business logic that the workflow activity is to execute.
        /// </summary>
        /// <param name="context">The context.</param>
        /// <remarks>
        /// For improved performance, Microsoft Dynamics CRM caches plug-in instances. 
        /// The plug-in's Execute method should be written to be stateless as the constructor 
        /// is not called for every invocation of the plug-in. Also, multiple system threads 
        /// could execute the plug-in at the same time. All per invocation state information 
        /// is stored in the context. This means that you should not use global variables in plug-ins.
        /// </remarks>
        protected override void Execute(CodeActivityContext context)
        {
            if (context == null)
            {
                throw new ArgumentNullException(nameof(context));
            }

            // Construct the Local plug-in context.
            LocalWorkflowContext localcontext = new LocalWorkflowContext(context);

            localcontext.Trace(string.Format(CultureInfo.InvariantCulture, "Entered {0}.Execute()", "Custom Workflow Activity"));

            try
            {
                ExecuteCrmWorkFlowActivity(context, localcontext);
            }
            catch (FaultException<OrganizationServiceFault> e)
            {
                localcontext.Trace(string.Format(CultureInfo.InvariantCulture, "Exception: {0}", e.ToString()));

                // Handle the exception.
                throw;
            }
            finally
            {
                localcontext.Trace(string.Format(CultureInfo.InvariantCulture, "Exiting {0}.Execute()", "Custom Workflow Activity"));
            }
        }

        protected virtual void ExecuteCrmWorkFlowActivity(CodeActivityContext context, LocalWorkflowContext localcontext)
        {
            // Do nothing. 
        }
    }
}