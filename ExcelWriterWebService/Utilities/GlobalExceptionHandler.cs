using ExcelWriterWebService.Models;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Web;
using System.Web.Http.ExceptionHandling;
using System.Web.Http.Results;

namespace ExcelWriterWebService.Utilities
{
	public class GlobalExceptionHandler : ExceptionHandler
	{
		public override void Handle(ExceptionHandlerContext context)
		{
			// 1. Access the original exception
			var exception = context.Exception;

			// 2. Perform your custom logic (e.g., call your existing logging/module code)
			// Log your exception here
			var db = new DatabaseContext();

			db.SaveToErrorLog(
				memberId: 0,
				exceptionType: exception.GetType().ToString(),
				exceptionMessage: exception.Message,
				exceptionSource: exception.Source,
				stackTrace: exception.StackTrace,
				innerExceptionType: exception.InnerException?.GetType().ToString(),
				innerExceptionMessage: exception.InnerException?.Message,
				innerExceptionSource: exception.InnerException?.Source,
				innerExceptionStackTrace: exception.InnerException?.StackTrace,
				context.Request.RequestUri.AbsoluteUri,
				context.Request.Headers.UserAgent.ToString(),
				GetClientIpAddress(context.Request),
				DateTime.Now
			);

			// 3. Create a custom JSON/HTTP response for the API client
			context.Result = new InternalServerErrorResult(context.Request);
		}

		public override bool ShouldHandle(ExceptionHandlerContext context)
		{
			// Ensure this handler processes the error
			return true;
		}
		private string GetClientIpAddress(HttpRequestMessage request)
		{
			if (request == null) return "Unknown";

			// Scenario A: Hosted on IIS / Web hosting
			if (request.Properties.TryGetValue("MS_HttpContext", out object contextObj))
			{
				var ctx = contextObj as HttpContextWrapper;
				if (ctx?.Request != null)
				{
					return ctx.Request.UserHostAddress;
				}
			}

			// Scenario B: Self-hosted / OWIN environment
			if (request.Properties.TryGetValue("MS_OwinContext", out object owinContextObj))
			{
				dynamic owinContext = owinContextObj;
				if (owinContext?.Request?.RemoteIpAddress != null)
				{
					return owinContext.Request.RemoteIpAddress;
				}
			}

			// Scenario C: Self-hosted directly via HttpSelfHostServer
			const string RemoteEndpointMessageProperty = "System.ServiceModel.Channels.RemoteEndpointMessageProperty";
			if (request.Properties.TryGetValue(RemoteEndpointMessageProperty, out object remoteEndpointObj))
			{
				dynamic remoteEndpoint = remoteEndpointObj;
				if (remoteEndpoint != null)
				{
					return remoteEndpoint.Address;
				}
			}

			return "Unknown";
		}
	}
}