namespace Topo.Services
{
    // Additional award id -> title. The same for everyone, so fetched once.
    public static class AdditionalAwardNames
    {
        private static Dictionary<string, string>? _names;

        public static async Task<Dictionary<string, string>> Get(ITerrainAPIService terrainAPIService)
        {
            if (_names == null)
            {
                var specifications = await terrainAPIService.GetAdditionalAwardSpecifications();
                _names = specifications.AwardDescriptions.GroupBy(a => a.id).ToDictionary(g => g.Key, g => g.First().title);
            }
            return _names;
        }
    }
}
