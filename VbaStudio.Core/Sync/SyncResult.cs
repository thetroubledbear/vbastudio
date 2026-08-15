using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace VbaStudio.Core.Sync;

public sealed record SyncResult(
    [property: JsonPropertyName("written")] IReadOnlyList<string> Written,
    [property: JsonPropertyName("deleted")] IReadOnlyList<string> Deleted,
    [property: JsonPropertyName("conflicts")] IReadOnlyList<string> Conflicts);
