using System;
using System.Collections.Generic;
using System.Net.WebSockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization; // Using standard .NET Framework serializer for simplicity

namespace WordOverlayProofreader.Addin
{
    public class SuggestionClient : IDisposable
    {
        private ClientWebSocket _ws;
        private const string ApiUrl = "wss://arabicdemo.abark.tech/ws/analyze";
        private const string AccessToken = "6c640104b56419ccfe17becd076521120445063cf4c9c6c30a60da45f89ec259fe74f46372279c505237f98db5230273fadee9fb5ea18924fa1a5b9f40b3dc92249933afc7e531a735216dddd9e1c17f9a5e9f515432e9a1a1c615c89924843acda8c5e570ff053a770bf467e348063d167ed90955946a000bf089ea5db683c5dbcbac2767049e4d7a57a6e2f074223624b822650677718e2d43c5055d12750c796bd9953ed810847356cf2e2ce70d0d8f1d0604946e90cc5cb649deeb7ff7a075d75e2a9bf3cb448312401a924bf413";
        private const string RefreshToken = "2ae62ce9f3c5834f126db18886ac7f4f965b80e4627b4d5d5b6a5f847360890126fe46e4b53eabf69588bcacfee06cfcaa8a8b4a3d0b1fefa19a4a5956714c63";

        public event EventHandler<List<Suggestion>> SuggestionsReceived;

        public async Task ScanAsync(string text)
        {
            try
            {
                Console.WriteLine($"[SuggestionClient] ScanAsync called with {text.Length} chars");
                if (_ws == null || _ws.State != WebSocketState.Open)
                {
                    Console.WriteLine($"[SuggestionClient] Connecting to {ApiUrl}...");
                    _ws = new ClientWebSocket();
                    _ws.Options.SetRequestHeader("Authorization", $"Bearer {AccessToken}");
                    await _ws.ConnectAsync(new Uri(ApiUrl), CancellationToken.None);
                    Console.WriteLine("[SuggestionClient] Connected to AI server");
                    _ = ReceiveLoop();
                }

                var request = new
                {
                    text = text,
                    requestId = Guid.NewGuid().ToString().Substring(0, 8),
                    accessToken = AccessToken,
                    refreshToken = RefreshToken
                };

                var serializer = new JavaScriptSerializer();
                var json = serializer.Serialize(request);
                Console.WriteLine($"[SuggestionClient] Sending request: {json.Substring(0, Math.Min(100, json.Length))}...");
                var bytes = Encoding.UTF8.GetBytes(json);
                await _ws.SendAsync(new ArraySegment<byte>(bytes), WebSocketMessageType.Text, true, CancellationToken.None);
                Console.WriteLine("[SuggestionClient] Request sent, waiting for response...");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[SuggestionClient] Error scanning text: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"Error scanning text: {ex.Message}");
                throw;
            }
        }

        private async Task ReceiveLoop()
        {
            var buffer = new byte[1024 * 64]; // Larger buffer for big responses
            var messageBuilder = new StringBuilder();
            
            try
            {
                while (_ws.State == WebSocketState.Open)
                {
                    var result = await _ws.ReceiveAsync(new ArraySegment<byte>(buffer), CancellationToken.None);
                    
                    if (result.MessageType == WebSocketMessageType.Close)
                    {
                        await _ws.CloseAsync(WebSocketCloseStatus.NormalClosure, string.Empty, CancellationToken.None);
                        break;
                    }
                    
                    if (result.MessageType == WebSocketMessageType.Text)
                    {
                        var chunk = Encoding.UTF8.GetString(buffer, 0, result.Count);
                        messageBuilder.Append(chunk);
                        
                        if (result.EndOfMessage)
                        {
                            var json = messageBuilder.ToString();
                            messageBuilder.Clear();
                            
                            try 
                            {
                                Console.WriteLine($"\n[SuggestionClient] ===== FULL JSON RESPONSE =====");
                                Console.WriteLine(json);
                                Console.WriteLine($"[SuggestionClient] ================================\n");
                                
                                var serializer = new JavaScriptSerializer();
                                var suggestions = serializer.Deserialize<List<Suggestion>>(json);
                                Console.WriteLine($"[SuggestionClient] Parsed {suggestions?.Count ?? 0} suggestions");
                                if (suggestions != null && suggestions.Count > 0)
                                {
                                    SuggestionsReceived?.Invoke(this, suggestions);
                                    Console.WriteLine($"[SuggestionClient] Invoked SuggestionsReceived event");
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"[SuggestionClient] Error parsing suggestions: {ex.Message}\nJSON: {json}");
                                System.Diagnostics.Debug.WriteLine($"Error parsing suggestions: {ex.Message}\nJSON: {json}");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ReceiveLoop error: {ex.Message}");
            }
        }

        public void Dispose()
        {
            _ws?.Dispose();
        }
    }

    public class Suggestion
    {
        public string type { get; set; }
        public string text { get; set; }
        public string suggestion { get; set; }
        public int from { get; set; }
        public int to { get; set; }
        public int occurence { get; set; }
        public string id { get; set; }
        public string requestId { get; set; }
    }
}
