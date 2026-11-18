namespace Topo.Model.Login
{
    // What "Remember Me" persists to browser local storage so a page refresh/reopen
    // doesn't require logging in again. Only written when the user opts in at login.
    public class PersistedSession
    {
        public string ClientId { get; set; } = string.Empty;
        public AuthenticationResult? AuthenticationResult { get; set; }
        public DateTime TokenExpiry { get; set; }
    }
}
