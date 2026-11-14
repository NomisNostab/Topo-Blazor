using Microsoft.AspNetCore.Components.Routing;
using Topo.Services;

namespace Topo.Shared
{
    // Who a menu group or link is shown to. Flags, so one entry can serve several audiences.
    [Flags]
    public enum NavAudience
    {
        None = 0,
        SignedOut = 1,
        Youth = 2,
        Leader = 4,
        SignedIn = Youth | Leader,
        Everyone = SignedOut | SignedIn
    }

    // One link in the menu. Icon is an Open Iconic class (oi-home). A link whose Href starts with http opens in a new tab.
    // ActivePrefix keeps the link highlighted on any page whose path starts with it (Scouting History has two pages).
    // TitleFor replaces Title when the text depends on state (Enable Logging / Stop Logging).
    public record NavItem(
        string Title,
        string Href,
        string Icon,
        NavAudience Audience = NavAudience.Everyone,
        NavLinkMatch Match = NavLinkMatch.Prefix,
        string? ActivePrefix = null,
        Func<StorageService, string>? TitleFor = null)
    {
        public bool IsExternal => Href.StartsWith("http", StringComparison.OrdinalIgnoreCase);

        public string TitleText(StorageService storage) => TitleFor?.Invoke(storage) ?? Title;
    }

    // A run of links. With a Title it shows as a collapsible heading (the title doubles as the key that remembers
    // whether it is open); without one the links are shown on their own with no heading.
    public record NavGroup(
        string? Title,
        NavAudience Audience,
        IReadOnlyList<NavItem> Items,
        bool OpenByDefault = true);

    // The menu, top to bottom. To add a page, add a NavItem to the group it belongs in; to change the order,
    // move the line (or move a whole group). Nothing else needs to change.
    public static class NavMenuItems
    {
        public const string HelpUrl = "https://github.com/NomisNostab/Topo-Blazor/wiki";

        public static readonly IReadOnlyList<NavGroup> Groups = new List<NavGroup>
        {
            new(null, NavAudience.Everyone, new NavItem[]
            {
                new("Home", "", "oi-home", Match: NavLinkMatch.All),
            }),
            new(null, NavAudience.SignedOut, new NavItem[]
            {
                new("Login", "login", "oi-lock-unlocked"),
                new("Help", HelpUrl, "oi-info"),
            }),
            new(null, NavAudience.Youth, new NavItem[]
            {
                new("Logbook", "logbook", "oi-book"),
                new("Additional Awards", "awards", "oi-shield"),
                new("Personal Progress", "progress/progress", "oi-person"),
            }),
            new("Badgework", NavAudience.Leader, new NavItem[]
            {
                new("OAS", "oas", "oi-compass"),
                new("SIA", "sia", "oi-clipboard"),
                new("Milestone", "milestone", "oi-flag"),
                new("Additional Awards", "awards", "oi-shield"),
            }),
            new("Reports", NavAudience.Leader, new NavItem[]
            {
                new("Members", "members", "oi-people", Match: NavLinkMatch.All),
                new("Program", "program", "oi-calendar", Match: NavLinkMatch.All),
                new("Logbook", "logbook", "oi-book"),
                new("Wallchart", "wallchart", "oi-bar-chart"),
                new("Approvals", "approvals/approvals", "oi-thumb-up"),
                new("Personal Progress", "progress/progress", "oi-person"),
                new("Scouting History", "scoutinghistory/member", "oi-clock", ActivePrefix: "scoutinghistory"),
            }),
            new("Tools", NavAudience.Leader, new NavItem[]
            {
                new("Patrol Planner", "members/patrolplanner", "oi-random"),
                new("Sign In Sheet", "signinsheet", "oi-pencil"),
                new("Program Ideas", "program/ideas", "oi-lightbulb"),
            }),
            new("Settings", NavAudience.SignedIn, new NavItem[]
            {
                new("Logout", "logout", "oi-lock-locked"),
                new("Help", HelpUrl, "oi-info"),
                new("Enable Logging", "logging", "oi-document", NavAudience.Leader,
                    TitleFor: storage => storage.LogToQueue ? "Stop Logging" : "Enable Logging"),
            }, OpenByDefault: false),
        };

        // The audience the current user belongs to
        public static NavAudience AudienceFor(StorageService storage) =>
            !storage.IsAuthenticated ? NavAudience.SignedOut
            : storage.IsYouthMember ? NavAudience.Youth
            : NavAudience.Leader;

        public static bool IsShownTo(NavAudience entry, NavAudience viewer) => (entry & viewer) != 0;
    }
}
