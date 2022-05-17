using Newtonsoft.Json;

namespace ISMExtensions.Extensions
{
    public static class ObjectExtension
    {
        public static string ToJson(this object value)
        {
            return JsonConvert.SerializeObject(value);
        }
    }
}
