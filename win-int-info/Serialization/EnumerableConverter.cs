using System;
using CsvHelper.TypeConversion;
using Newtonsoft.Json;

namespace win_int_info.Serialization
{
    internal class EnumerableConverter : ITypeConverter
    {
        public string ConvertToString(TypeConverterOptions options, object value)
        {
            JsonSerializerSettings settings = new JsonSerializerSettings {Culture = options.CultureInfo};
            return JsonConvert.SerializeObject(value,Formatting.Indented, settings);
        }

        public object ConvertFromString(TypeConverterOptions options, string text)
        {
            JsonSerializerSettings settings = new JsonSerializerSettings { Culture = options.CultureInfo };
            return JsonConvert.DeserializeObject(text, settings);
        }

        public bool CanConvertFrom(Type type)
        {
            throw new NotImplementedException();
        }

        public bool CanConvertTo(Type type)
        {
            if (type == typeof(string))
                return true;
            return false;
        }
    }
}