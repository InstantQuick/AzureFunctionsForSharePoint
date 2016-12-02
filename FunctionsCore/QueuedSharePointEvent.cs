namespace FunctionsCore
{
    public class QueuedSharePointEvent : QueuedFunctionEvent
    {
        public string ClientId { get; set; }
        public string AppWebUrl { get; set; }
        public string AppAccessToken { get; set; }
        public string UserAccessToken { get; set; }
    }
}
