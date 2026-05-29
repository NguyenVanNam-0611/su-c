using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Diff_tool.Services;

namespace Diff_tool.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    [Authorize]
    public class CompareController : ControllerBase
    {
        private readonly IHttpClientFactory _httpClientFactory;
        private readonly IConfiguration _config;
        private readonly ChecksheetWordService _checksheetWordService;
        private readonly ILogger<CompareController> _logger;

        public CompareController(
            IHttpClientFactory httpClientFactory,
            IConfiguration config,
            ChecksheetWordService checksheetWordService,
            ILogger<CompareController> logger)
        {
            _httpClientFactory = httpClientFactory;
            _config = config;
            _checksheetWordService = checksheetWordService;
            _logger = logger;
        }

        [HttpGet("status/{jobId}")]
        public async Task<IActionResult> GetStatus(string jobId)
        {
            var fastApiUrl = _config["FastApi:BaseUrl"] ?? "http://localhost:8000";
            var client = _httpClientFactory.CreateClient();
            var response = await client.GetAsync($"{fastApiUrl}/api/status/job/{jobId}");
            var content = await response.Content.ReadAsStringAsync();
            if (!response.IsSuccessStatusCode)
                return StatusCode((int)response.StatusCode, new { message = content });
            return Content(content, "application/json");
        }

        [HttpPost("export")]
        public IActionResult Export([FromBody] ChecksheetRequest request)
        {
            try
            {
                var sections = request?.Sections ?? new();

                foreach (var s in sections)
                {
                    s.Changes = (s.Changes ?? new())
                        .Where(c =>
                        {
                            var kind = c.ChangeKind?.ToLower();
                            bool keep = kind is "replace" or "insert" or "delete";

                            _logger.LogInformation(
                                "change type={Type} kind={Kind} => {Action}",
                                c.Type, c.ChangeKind, keep ? "KEEP" : "SKIP");

                            if (keep)
                            {
                                _logger.LogInformation(
                                    "  left={LeftPreview} page={LeftPage} nodeNull={LeftNull} | right={RightPreview} page={RightPage} nodeNull={RightNull}",
                                    c.Left?.PreviewText, c.Left?.Page, c.Left?.Node == null,
                                    c.Right?.PreviewText, c.Right?.Page, c.Right?.Node == null);
                                _logger.LogInformation(
                                    "  wordDiff null={WdNull} spans={SpanCount}",
                                    c.WordDiff == null, c.WordDiff?.Spans?.Count ?? -1);

                                if (c.WordDiff?.Spans != null)
                                    foreach (var sp in c.WordDiff.Spans)
                                        _logger.LogInformation(
                                            "    span type={SpType} text={Text} old={Old} new={New}",
                                            sp.Type, sp.Text, sp.OldText, sp.NewText);
                            }

                            return keep;
                        })
                        .ToList();
                }

                var fileBytes = _checksheetWordService.GenerateChecksheet(sections);
                var fileName = $"Checksheet_{request?.JobId ?? "export"}.docx";

                return File(
                    fileBytes,
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    fileName);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Export thất bại");
                return StatusCode(500, new
                {
                    message = ex.Message,
                    type = ex.GetType().FullName,
                    stackTrace = ex.StackTrace,
                    inner = ex.InnerException?.Message
                });
            }
        }
    }
}