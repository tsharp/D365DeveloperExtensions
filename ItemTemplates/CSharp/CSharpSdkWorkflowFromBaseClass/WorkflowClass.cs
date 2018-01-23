using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;

namespace $rootnamespace$
{
    public class $safeitemname$ : WorkFlowActivityBase
    {
        protected override void ExecuteCrmWorkFlowActivity(CodeActivityContext context, LocalWorkflowContext localContext)
        {
            if (context == null)
                throw new ArgumentNullException(nameof(context));
            if (localContext == null)
                throw new ArgumentNullException(nameof(localContext));

            // TODO: Implement your custom code
        }
    }
}