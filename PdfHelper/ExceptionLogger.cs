using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using log4net;
using System.Reflection;

namespace PdfFunctions
{
    public static class ExceptionLogger
    {
        private static readonly ILog Log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);
        public static void logException(Exception ex)
        {
            var sb = new StringBuilder(ex.ToString());
            sb.AppendLine("--TARGET SITE--");
            sb.AppendLine(ex.TargetSite.Name);
            sb.AppendLine("--INNER EXCEPTION");
            sb.AppendLine(ex.InnerException != null ? ex.InnerException.ToString() : "N/A");
            sb.AppendLine("--SOURCE--");
            sb.AppendLine(ex.Source);
            sb.AppendLine("--STACK TRACE--");
            sb.AppendLine(ex.StackTrace);
            Log.Error(sb.ToString(), ex);
        }
    }
}
