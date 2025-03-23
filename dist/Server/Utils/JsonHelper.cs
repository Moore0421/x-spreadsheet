using System;
using System.Text.Json;

namespace SpreadsheetServer.Utils
{
    public static class JsonHelper
    {
        // 深度比较两个对象是否相同
        public static bool IsEqual(JsonElement obj1, JsonElement obj2)
        {
            return JsonSerializer.Serialize(obj1) == JsonSerializer.Serialize(obj2);
        }
    }
}