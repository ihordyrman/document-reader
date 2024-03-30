using System.Text.Json.Serialization;

namespace DocumentReader;

[JsonSerializable(typeof(Config))]
internal partial class ConfigContext : JsonSerializerContext;
