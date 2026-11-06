using System.Text.Json;

namespace Topo.Model.SiaProgress
{
    public static class SiaTemplateParser
    {
        // The questions of one of Terrain's SIA templates in the order they are asked, and the wording of the options a
        // Scout picks from. A template that can't be read gives no questions.
        public static SiaTemplate Parse(string json)
        {
            var template = new SiaTemplate();
            try
            {
                using var document = JsonDocument.Parse(json);
                if (document.RootElement.ValueKind == JsonValueKind.Object && document.RootElement.TryGetProperty("document", out var body))
                    Walk(body, template);
            }
            catch (JsonException)
            {
            }
            return template;
        }

        private static void Walk(JsonElement element, SiaTemplate template)
        {
            if (element.ValueKind == JsonValueKind.Array)
            {
                foreach (var item in element.EnumerateArray())
                    Walk(item, template);
            }
            else if (element.ValueKind == JsonValueKind.Object)
            {
                var id = Text(element, "id");
                if (id != "" && element.TryGetProperty("label", out _) && element.TryGetProperty("type", out _))
                {
                    // a question; the file uploader is shown as the files attached
                    if (Text(element, "type") != "file_uploader")
                        template.Questions.Add(new SiaQuestion { Id = id, Label = Text(element, "label").TrimEnd('*').Trim() });
                }
                else if (id != "" && element.TryGetProperty("title", out _))
                {
                    template.Options[id] = Text(element, "title");
                }

                foreach (var property in element.EnumerateObject())
                    Walk(property.Value, template);
            }
        }

        private static string Text(JsonElement element, string name) =>
            element.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.String ? value.GetString() ?? "" : "";
    }
}
