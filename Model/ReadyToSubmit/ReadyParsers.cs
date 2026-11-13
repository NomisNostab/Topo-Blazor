using System.Globalization;
using System.Text.Json;

namespace Topo.Model.ReadyToSubmit
{
    // Reads Terrain's responses into the simple shapes the report works from
    public static class ReadyParsers
    {
        // Achievements of one type, as Terrain returns them for a member or for the whole unit
        public static List<AchievementRecord> ParseAchievements(string json)
        {
            var records = new List<AchievementRecord>();
            try
            {
                using var document = JsonDocument.Parse(json);
                if (document.RootElement.ValueKind != JsonValueKind.Object || !document.RootElement.TryGetProperty("results", out var results) || results.ValueKind != JsonValueKind.Array)
                    return records;
                foreach (var item in results.EnumerateArray())
                {
                    var record = new AchievementRecord
                    {
                        Id = Text(item, "id"),
                        MemberId = Text(item, "member_id"),
                        Type = Text(item, "type"),
                        Section = Text(item, "section").Trim().ToLowerInvariant(),
                        Template = Text(item, "template"),
                        Status = Text(item, "status"),
                        RequirementStatus = Text(item, "milestone_requirement_status"),
                        Answers = Flatten(item, "answers"),
                        Meta = Flatten(item, "achievement_meta")
                    };
                    if (DateTime.TryParse(Text(item, "status_updated"), CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal, out var updated))
                        record.StatusUpdated = updated.ToLocalTime();
                    records.Add(record);
                }
            }
            catch (JsonException)
            {
            }
            return records;
        }

        // The questions of a form template in the order they are asked
        public static List<FormQuestion> ParseFormQuestions(string json)
        {
            var questions = new List<FormQuestion>();
            try
            {
                using var document = JsonDocument.Parse(json);
                if (document.RootElement.ValueKind == JsonValueKind.Object && document.RootElement.TryGetProperty("document", out var body))
                    WalkQuestions(body, questions);
            }
            catch (JsonException)
            {
            }
            return questions;
        }

        // Everything that has to be ticked off in an OAS stage's template, in the order they are asked
        public static List<FormQuestion> ParseOasInputs(string json)
        {
            var ids = new List<FormQuestion>();
            try
            {
                using var document = JsonDocument.Parse(json);
                if (document.RootElement.ValueKind != JsonValueKind.Object || !document.RootElement.TryGetProperty("document", out var body) || body.ValueKind != JsonValueKind.Array)
                    return ids;
                foreach (var page in body.EnumerateArray())
                {
                    if (page.ValueKind != JsonValueKind.Object || !page.TryGetProperty("input_groups", out var groups) || groups.ValueKind != JsonValueKind.Array)
                        continue;
                    foreach (var group in groups.EnumerateArray())
                    {
                        if (group.ValueKind != JsonValueKind.Object || !group.TryGetProperty("inputs", out var inputs) || inputs.ValueKind != JsonValueKind.Array)
                            continue;
                        foreach (var input in inputs.EnumerateArray())
                        {
                            var id = Text(input, "id");
                            if (id != "" && id != "file_uploader")
                                ids.Add(new FormQuestion { Id = id, Type = Text(input, "type"), Label = Text(input, "label").Trim() });
                        }
                    }
                }
            }
            catch (JsonException)
            {
            }
            return ids;
        }

        private static void WalkQuestions(JsonElement element, List<FormQuestion> questions)
        {
            if (element.ValueKind == JsonValueKind.Array)
            {
                foreach (var item in element.EnumerateArray())
                    WalkQuestions(item, questions);
            }
            else if (element.ValueKind == JsonValueKind.Object)
            {
                var id = Text(element, "id");
                if (id != "" && element.TryGetProperty("type", out _) && element.TryGetProperty("label", out _))
                    questions.Add(new FormQuestion { Id = id, Type = Text(element, "type"), Label = Text(element, "label").TrimEnd('*').Trim() });
                foreach (var property in element.EnumerateObject())
                    WalkQuestions(property.Value, questions);
            }
        }

        // An object's values as text: arrays are joined, true and false kept as words, and a nested object (such as a date range)
        // kept as its own text so that having one counts as an answer
        private static Dictionary<string, string> Flatten(JsonElement parent, string name)
        {
            var values = new Dictionary<string, string>();
            if (!parent.TryGetProperty(name, out var obj) || obj.ValueKind != JsonValueKind.Object)
                return values;
            foreach (var property in obj.EnumerateObject())
            {
                switch (property.Value.ValueKind)
                {
                    case JsonValueKind.String:
                        values[property.Name] = property.Value.GetString() ?? "";
                        break;
                    case JsonValueKind.Number:
                        values[property.Name] = property.Value.GetRawText();
                        break;
                    case JsonValueKind.True:
                        values[property.Name] = "true";
                        break;
                    case JsonValueKind.False:
                        values[property.Name] = "false";
                        break;
                    case JsonValueKind.Object:
                        values[property.Name] = property.Value.GetRawText() == "{}" ? "" : property.Value.GetRawText();
                        break;
                    case JsonValueKind.Array:
                        values[property.Name] = string.Join(", ", property.Value.EnumerateArray().Select(a => a.ValueKind == JsonValueKind.String ? a.GetString() ?? "" : a.GetRawText()));
                        break;
                }
            }
            return values;
        }

        private static string Text(JsonElement element, string name) =>
            element.ValueKind == JsonValueKind.Object && element.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.String ? value.GetString() ?? "" : "";
    }
}
