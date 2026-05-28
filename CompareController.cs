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
        private readonly ChecksheetService _checksheetService;
        private readonly ChecksheetWordService _checksheetWordService;

        public CompareController(
            IHttpClientFactory httpClientFactory,
            IConfiguration config,
            ChecksheetService checksheetService,
            ChecksheetWordService checksheetWordService)
        {
            _httpClientFactory = httpClientFactory;
            _config = config;
            _checksheetService = checksheetService;
            _checksheetWordService = checksheetWordService;
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

        // ── Excel: tạm thời vô hiệu hóa ──────────────────────────────────────
        // [HttpPost("export")]
        // public IActionResult Export([FromBody] ChecksheetRequest request) { ... }

        [HttpPost("export")]
        public IActionResult Export([FromBody] ChecksheetRequest _)
        {
            return StatusCode(503, new { message = "Export Excel đã bị vô hiệu hóa. Vui lòng dùng export-word." });
        }

        // ── Word export ───────────────────────────────────────────────────────
        [HttpPost("export-word")]
        public IActionResult ExportWord([FromBody] ChecksheetRequest request)
        {
            try
            {
                foreach (var s in request?.Sections ?? new())
                    foreach (var c in s.Changes ?? new())
                    {
                        Console.WriteLine($"change type={c.Type} kind={c.ChangeKind}");
                        Console.WriteLine($"  left:  preview={c.Left?.PreviewText}, nodeNull={c.Left?.Node == null}");
                        Console.WriteLine($"  right: preview={c.Right?.PreviewText}, nodeNull={c.Right?.Node == null}");
                        Console.WriteLine($"  wordDiff null={c.WordDiff == null}, spans={c.WordDiff?.Spans?.Count ?? -1}");
                        if (c.WordDiff?.Spans != null)
                            foreach (var sp in c.WordDiff.Spans)
                                Console.WriteLine($"    span: type={sp.Type} text={sp.Text} old={sp.OldText} new={sp.NewText}");
                    }

                var fileBytes = _checksheetWordService.GenerateChecksheet(request?.Sections ?? new());
                var fileName = $"Checksheet_{request?.JobId ?? "export"}.docx";
                return File(fileBytes,
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    fileName);
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { message = ex.Message });
            }
        }
    }
}