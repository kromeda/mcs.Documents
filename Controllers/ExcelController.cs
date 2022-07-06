using Documents.Reports;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Models;

namespace Documents.Controllers
{
    [ApiController]
    [Route("excel")]
    public class ExcelController : ControllerBase
    {
        private const string sentReportName = "Отчет направленных уведомлений";
        private readonly ILogger<ExcelController> logger;

        public ExcelController(ILogger<ExcelController> logger)
        {
            this.logger = logger;
        }

        [HttpPost("fl/notifications/sent")]
        public ActionResult<FileDocumentView> SentIndividualNotificationsReport(SentIndividualArgs args)
        {
            var report = new SentIndividualNotifications();
            var file = report.Export(args);
            var document = new FileDocumentView() { Data = file, Name = sentReportName + ".xlsx" };
            logger.LogInformation((int)LogType.FileGenerated, "Сформирован файл: {FileName}.", sentReportName);

            return Ok(document);
        }
    }
}