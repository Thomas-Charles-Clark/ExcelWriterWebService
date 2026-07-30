using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.SqlClient;
using System.Data;
using System.Data.Linq;
using System.IO;
using System.Linq;
using System.Web;
using ExcelWriterWebService.Utilities;
using System.Threading.Tasks;
using System.Data.Linq.Mapping;
using System.Reflection;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Configuration;

namespace ExcelWriterWebService.Models
{
	public partial class DatabaseContext : System.Data.Linq.DataContext
	{

		private static System.Data.Linq.Mapping.MappingSource mappingSource = new AttributeMappingSource();

		#region Extensibility Method Definitions
		partial void OnCreated();
		#endregion

		public DatabaseContext() :
				base(ConfigurationManager.ConnectionStrings["FlockConnectionString"].ToString(), mappingSource)
		{
			OnCreated();
		}

		public DatabaseContext(string connection) :
				base(connection, mappingSource)
		{
			OnCreated();
		}

		public DatabaseContext(System.Data.IDbConnection connection) :
				base(connection, mappingSource)
		{
			OnCreated();
		}

		public DatabaseContext(string connection, System.Data.Linq.Mapping.MappingSource mappingSource) :
				base(connection, mappingSource)
		{
			OnCreated();
		}

		public DatabaseContext(System.Data.IDbConnection connection, System.Data.Linq.Mapping.MappingSource mappingSource) :
				base(connection, mappingSource)
		{
			OnCreated();
		}

		[Function(Name = "pub.LogError")]
		public ISingleResult<ErrorLogResult> SaveToErrorLog(
			[Parameter(Name = "MemberID", DbType = "Int")] int? memberId,
			[Parameter(Name = "ExceptionType", DbType = "NVarChar(MAX)")] string exceptionType,
			[Parameter(Name = "ExceptionMessage", DbType = "NVarChar(MAX)")] string exceptionMessage,
			[Parameter(Name = "ExceptionSource", DbType = "NVarChar(MAX)")] string exceptionSource,
			[Parameter(Name = "StackTrace", DbType = "NVarChar(MAX)")] string stackTrace,
			[Parameter(Name = "InnerExceptionType", DbType = "NVarChar(MAX)")] string innerExceptionType,
			[Parameter(Name = "InnerExceptionMessage", DbType = "NVarChar(MAX)")] string innerExceptionMessage,
			[Parameter(Name = "InnerExceptionSource", DbType = "NVarChar(MAX)")] string innerExceptionSource,
			[Parameter(Name = "InnerExceptionStackTrace", DbType = "NVarChar(MAX)")] string innerExceptionStackTrace,
			[Parameter(Name = "URL", DbType = "NVarChar(500)")] string url,
			[Parameter(Name = "UserAgent", DbType = "NVarChar(255)")] string userAgent,
			[Parameter(Name = "IP", DbType = "NVarChar(50)")] string ip,
			[Parameter(Name = "ErrorDT", DbType = "DateTime2(2)")] DateTime errorDt,
			[Parameter(Name = "PlatformSwitch", DbType = "TinyInt")] byte platformSwitch = 9)
		{
			IExecuteResult result = this.ExecuteMethodCall(this, ((MethodInfo)(MethodInfo.GetCurrentMethod())), 
				memberId,
				exceptionType,
				exceptionMessage,
				exceptionSource,
				stackTrace,
				innerExceptionType,
				innerExceptionMessage,
				innerExceptionSource,
				innerExceptionStackTrace,
				url,
				userAgent,
				ip,
				errorDt,
				platformSwitch);

			return ((ISingleResult<ErrorLogResult>)(result.ReturnValue));
		}
	}
}