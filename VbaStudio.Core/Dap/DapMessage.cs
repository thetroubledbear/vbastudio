using System.Text.Json;
using System.Text.Json.Serialization;

namespace VbaStudio.Core.Dap;

public sealed record DapRequest(
    [property: JsonPropertyName("seq")] int Seq,
    [property: JsonPropertyName("command")] string Command,
    [property: JsonPropertyName("arguments")] JsonElement? Arguments)
{
    [JsonPropertyName("type")]
    public string Type => "request";
}

public sealed record DapResponse(
    [property: JsonPropertyName("seq")] int Seq,
    [property: JsonPropertyName("request_seq")] int RequestSeq,
    [property: JsonPropertyName("success")] bool Success,
    [property: JsonPropertyName("command")] string Command,
    [property: JsonPropertyName("body")] object? Body,
    [property: JsonPropertyName("message")] string? Message = null)
{
    [JsonPropertyName("type")]
    public string Type => "response";
}

public sealed record DapEvent(
    [property: JsonPropertyName("seq")] int Seq,
    [property: JsonPropertyName("event")] string Event,
    [property: JsonPropertyName("body")] object? Body)
{
    [JsonPropertyName("type")]
    public string Type => "event";
}
