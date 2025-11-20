using System.Net;
using System.Net.WebSockets;
using System.Text;
using System.Text.Json;

namespace SuggestionServer
{
    public class Program
    {
        public static async Task Main(string[] args)
        {
            var httpListener = new HttpListener();
            httpListener.Prefixes.Add("http://localhost:8080/");
            httpListener.Start();
            Console.WriteLine("Suggestion Server started on ws://localhost:8080/");

            while (true)
            {
                var context = await httpListener.GetContextAsync();
                if (context.Request.IsWebSocketRequest)
                {
                    ProcessRequest(context);
                }
                else
                {
                    context.Response.StatusCode = 400;
                    context.Response.Close();
                }
            }
        }

        private static async void ProcessRequest(HttpListenerContext context)
        {
            WebSocketContext wsContext = null;
            try
            {
                wsContext = await context.AcceptWebSocketAsync(null);
                Console.WriteLine("Client connected");
                var webSocket = wsContext.WebSocket;

                var buffer = new byte[1024 * 4];
                while (webSocket.State == WebSocketState.Open)
                {
                    var result = await webSocket.ReceiveAsync(new ArraySegment<byte>(buffer), CancellationToken.None);
                    if (result.MessageType == WebSocketMessageType.Close)
                    {
                        await webSocket.CloseAsync(WebSocketCloseStatus.NormalClosure, string.Empty, CancellationToken.None);
                    }
                    else if (result.MessageType == WebSocketMessageType.Text)
                    {
                        var json = Encoding.UTF8.GetString(buffer, 0, result.Count);
                        Console.WriteLine($"Received: {json}");
                        
                        // Parse request
                        try 
                        {
                            var request = JsonSerializer.Deserialize<ScanRequest>(json);
                            if (request != null && request.type == "scan")
                            {
                                await SendMockSuggestions(webSocket, request);
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error parsing request: {ex.Message}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"WebSocket error: {ex.Message}");
            }
            finally
            {
                if (wsContext != null)
                    wsContext.WebSocket.Dispose();
            }
        }

        private static async Task SendMockSuggestions(WebSocket webSocket, ScanRequest request)
        {
            // Mock logic: generate random suggestions based on input text length or content
            var suggestions = new List<Suggestion>();
            
            // Always add a few demo suggestions
            suggestions.Add(new Suggestion
            {
                type = "spelling",
                text = "teh",
                suggestion = "the",
                from = 0,
                to = 3,
                occurence = 0,
                id = Guid.NewGuid().ToString(),
                requestId = request.requestId
            });

            suggestions.Add(new Suggestion
            {
                type = "grammar",
                text = "is",
                suggestion = "are",
                from = 5,
                to = 7,
                occurence = 0,
                id = Guid.NewGuid().ToString(),
                requestId = request.requestId
            });

            // If text contains specific keywords, add more specific types
            if (request.text.Contains("style"))
            {
                 suggestions.Add(new Suggestion
                {
                    type = "style",
                    text = "utilize",
                    suggestion = "use",
                    from = request.text.IndexOf("style"), // Just a dummy position
                    to = request.text.IndexOf("style") + 5,
                    occurence = 0,
                    id = Guid.NewGuid().ToString(),
                    requestId = request.requestId
                });
            }

            // Send as a single batch for simplicity, or simulate streaming
            var response = JsonSerializer.Serialize(suggestions);
            var bytes = Encoding.UTF8.GetBytes(response);
            await webSocket.SendAsync(new ArraySegment<byte>(bytes), WebSocketMessageType.Text, true, CancellationToken.None);
            Console.WriteLine($"Sent {suggestions.Count} suggestions");
        }
    }

    public class ScanRequest
    {
        public string type { get; set; }
        public string requestId { get; set; }
        public string documentId { get; set; }
        public Range range { get; set; }
        public string text { get; set; }
    }

    public class Range
    {
        public int start { get; set; }
        public int end { get; set; }
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
