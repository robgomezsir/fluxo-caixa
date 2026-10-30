using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json.Serialization;

namespace fluxo_caixa_crud.Models
{
    public static class ApiJson
    {
        public static readonly JsonSerializerSettings Settings = new JsonSerializerSettings
        {
            ContractResolver = new CamelCasePropertyNamesContractResolver(),
            NullValueHandling = NullValueHandling.Ignore
        };

        public static string Serialize(object value)
        {
            return JsonConvert.SerializeObject(value, Settings);
        }
    }

    public class ApiRequest
    {
        public string Id { get; set; }
        public string Action { get; set; }
        public string Entity { get; set; }
        public JObject Payload { get; set; }
    }

    public class ApiResponse
    {
        public string Id { get; set; }
        public bool Success { get; set; }
        public object Data { get; set; }
        public string Error { get; set; }

        public static ApiResponse Ok(string id, object data)
        {
            return new ApiResponse { Id = id, Success = true, Data = data };
        }

        public static ApiResponse Fail(string id, string error)
        {
            return new ApiResponse { Id = id, Success = false, Error = error };
        }
    }
}
