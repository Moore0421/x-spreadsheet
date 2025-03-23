using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;

namespace SpreadsheetServer.Controllers
{
    [ApiController]
    [Route("api")]
    public class SpreadsheetController : ControllerBase
    {
        private readonly string DATA_DIR = Path.Combine(Directory.GetCurrentDirectory(), "data");
        private readonly string TEMP_DIR = Path.Combine(Directory.GetCurrentDirectory(), "temp");
        private readonly string UPLOADS_DIR = Path.Combine(Directory.GetCurrentDirectory(), "uploads");
        private readonly JsonSerializerOptions _jsonOptions = new JsonSerializerOptions { WriteIndented = true };

        // 获取文件名
        private string GetFileName(string id, bool pure)
        {
            string fileName = pure
                ? $"spreadsheet-data-pure-{id}.json"
                : $"spreadsheet-data-{id}.json";
            return Path.Combine(DATA_DIR, fileName);
        }

        // 分片上传接口
        [HttpPost("chunk-upload")]
        public ActionResult ChunkUpload([FromQuery] string chunkIndex, [FromQuery] string totalChunks, [FromQuery] string id)
        {
            try
            {
                // 读取请求体
                string requestBody;
                using (var reader = new StreamReader(Request.Body))
                {
                    requestBody = reader.ReadToEndAsync().Result;
                }

                // 确保临时目录存在
                if (!Directory.Exists(TEMP_DIR))
                {
                    Directory.CreateDirectory(TEMP_DIR);
                }

                string chunkPath = Path.Combine(TEMP_DIR, $"{id}-{chunkIndex}");
                System.IO.File.WriteAllText(chunkPath, requestBody);

                return Ok(new
                {
                    success = true,
                    chunkIndex,
                    message = $"Chunk {chunkIndex} of {totalChunks} received"
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { success = false, error = ex.Message });
            }
        }

        // 合并分片接口
        [HttpPost("merge-chunks")]
        public ActionResult MergeChunks([FromQuery] string id, [FromQuery] string totalChunks, [FromQuery] string pure)
        {
            try
            {
                int chunksCount = int.Parse(totalChunks);
                string finalPath = GetFileName(id, pure == "true");
                string rawData = "";

                Console.WriteLine($"开始合并，总分片数: {chunksCount}");

                // 按顺序读取并合并所有分片
                for (int i = 0; i < chunksCount; i++)
                {
                    string chunkPath = Path.Combine(TEMP_DIR, $"{id}-{i}");
                    string chunkContent = System.IO.File.ReadAllText(chunkPath);
                    
                    // 解析分片内容
                    var chunkData = JsonSerializer.Deserialize<Dictionary<string, object>>(chunkContent);
                    
                    // 提取分片中的实际数据
                    if (chunkData != null && chunkData.ContainsKey("data"))
                    {
                        rawData += chunkData["data"].ToString();
                    }

                    // 删除分片文件
                    System.IO.File.Delete(chunkPath);
                }

                // 解析完整的JSON字符串
                var finalData = JsonSerializer.Deserialize<JsonElement>(rawData);
                Console.WriteLine($"合并后数据对象类型: {finalData.GetType().Name}");

                // 保存为格式化的JSON
                System.IO.File.WriteAllText(finalPath, JsonSerializer.Serialize(finalData, _jsonOptions));
                Console.WriteLine("最终文件已保存");

                return Ok(new { success = true, message = "File merged successfully" });
            }
            catch (Exception ex)
            {
                Console.WriteLine($"合并错误: {ex}");
                return StatusCode(500, new { success = false, error = ex.Message });
            }
        }

        // 获取表格总数接口
        [HttpGet("getSheetCount")]
        public ActionResult GetSheetCount([FromQuery] string id)
        {
            try
            {
                string fileName = GetFileName(id, false);
                Console.WriteLine($"fileName {fileName}");

                if (System.IO.File.Exists(fileName))
                {
                    string content = System.IO.File.ReadAllText(fileName);
                    var data = JsonSerializer.Deserialize<JsonElement[]>(content);
                    return Ok(new
                    {
                        success = true,
                        count = data.Length
                    });
                }
                else
                {
                    return Ok(new { success = false });
                }
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { success = false, error = ex.Message });
            }
        }

        // 获取单个表格数据接口
        [HttpGet("getSheet")]
        public ActionResult GetSheet([FromQuery] string id, [FromQuery] int index)
        {
            try
            {
                string fileName = GetFileName(id, false);

                if (System.IO.File.Exists(fileName))
                {
                    string content = System.IO.File.ReadAllText(fileName);
                    var data = JsonSerializer.Deserialize<JsonElement[]>(content);
                    
                    if (index >= 0 && index < data.Length)
                    {
                        return Ok(new
                        {
                            success = true,
                            sheet = data[index]
                        });
                    }
                    else
                    {
                        return Ok(new { success = false, message = "Sheet index out of range" });
                    }
                }
                else
                {
                    return Ok(new {});
                }
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { success = false, error = ex.Message });
            }
        }

        // 同步纯数据到完整文件接口
        [HttpPost("sync-data")]
        public async Task<ActionResult> SyncData([FromQuery] string id, IFormFile pureDataFile)
        {
            if (pureDataFile == null)
            {
                return BadRequest(new { success = false, message = "No file uploaded" });
            }

            string fullFileName = GetFileName(id, false);
            string tempFilePath = Path.Combine(UPLOADS_DIR, Guid.NewGuid().ToString() + ".json");

            try
            {
                // 检查完整数据文件是否存在
                if (!System.IO.File.Exists(fullFileName))
                {
                    return NotFound(new
                    {
                        success = false,
                        message = "Full data file not found"
                    });
                }

                // 保存上传的文件
                using (var stream = new FileStream(tempFilePath, FileMode.Create))
                {
                    await pureDataFile.CopyToAsync(stream);
                }

                // 读取上传的纯数据文件
                string pureDataContent = System.IO.File.ReadAllText(tempFilePath);
                var pureData = JsonSerializer.Deserialize<JsonElement[]>(pureDataContent);

                // 读取完整数据文件
                string fullDataContent = System.IO.File.ReadAllText(fullFileName);
                var fullData = JsonSerializer.Deserialize<List<JsonElement>>(fullDataContent);

                int modifiedSheets = 0;
                var fullDataDict = new Dictionary<string, JsonElement>();

                // 将完整数据转换为字典，以便更容易查找
                for (int i = 0; i < fullData.Count; i++)
                {
                    var element = fullData[i];
                    if (element.TryGetProperty("name", out var nameElement))
                    {
                        string name = nameElement.GetString();
                        fullDataDict[name] = element;
                    }
                }

                // 处理合并逻辑（简化版 - 在实际应用中需要更复杂的合并逻辑）
                // 注意：这里的实现是简化的，不包含原始JavaScript代码中的复杂递归对象更新
                bool hasChanges = false;
                foreach (var pureSheet in pureData)
                {
                    if (pureSheet.TryGetProperty("name", out var nameElement))
                    {
                        string name = nameElement.GetString();
                        if (fullDataDict.ContainsKey(name))
                        {
                            // 这里只是简单替换，实际应用中应使用更复杂的合并逻辑
                            fullDataDict[name] = pureSheet;
                            modifiedSheets++;
                            hasChanges = true;
                        }
                    }
                }

                // 删除临时上传的文件
                System.IO.File.Delete(tempFilePath);

                // 只有在有修改时才保存文件
                if (hasChanges)
                {
                    // 将字典转换回列表并保存
                    var updatedFullData = fullDataDict.Values.ToList();
                    System.IO.File.WriteAllText(fullFileName, JsonSerializer.Serialize(updatedFullData, _jsonOptions));
                    
                    return Ok(new
                    {
                        success = true,
                        message = $"Successfully synchronized data. Modified {modifiedSheets} sheets."
                    });
                }
                else
                {
                    return Ok(new
                    {
                        success = true,
                        message = "No changes needed, data already in sync."
                    });
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"同步数据错误: {ex}");
                // 确保清理临时文件
                if (System.IO.File.Exists(tempFilePath))
                {
                    System.IO.File.Delete(tempFilePath);
                }
                return StatusCode(500, new { success = false, error = ex.Message });
            }
        }
    }
}