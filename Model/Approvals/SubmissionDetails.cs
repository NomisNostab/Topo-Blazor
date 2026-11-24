using System.Text.Json;

namespace Topo.Model.Approvals
{
    // What Terrain's "View submission details" shows for a request: the form the youth member filled in (the template's
    // sections, questions and ticks) with their answers, and the files they attached. Terrain makes it by reading the
    // achievement record (answers, uploads, which template and version it was filled in on) and the template file.
    public class SubmissionDetails
    {
        public List<DetailSection> Sections { get; init; } = new();
        public List<DetailUpload> Uploads { get; init; } = new();

        // The achievement record is Terrain's JSON for GET members/{id}/achievements/{achievementId}; the template is
        // the template file it names. Returns null when either cannot be read.
        public static SubmissionDetails? Build(string? templateJson, string? achievementJson)
        {
            if (!TryParse(templateJson, out var templateDoc) || !TryParse(achievementJson, out var achievementDoc))
            {
                return null;
            }
            using (templateDoc)
            using (achievementDoc)
            {
                var achievement = achievementDoc!.RootElement;
                if (templateDoc!.RootElement.ValueKind != JsonValueKind.Object || achievement.ValueKind != JsonValueKind.Object
                    || !templateDoc.RootElement.TryGetProperty("document", out var document) || document.ValueKind != JsonValueKind.Array)
                {
                    return null;
                }

                var answers = achievement.TryGetProperty("answers", out var a) && a.ValueKind == JsonValueKind.Object ? a : (JsonElement?)null;
                var uploads = ReadUploads(achievement);
                var details = new SubmissionDetails { Uploads = uploads };

                foreach (var section in document.EnumerateArray())
                {
                    var detailSection = new DetailSection { Title = Text(section, "title") };
                    if (section.TryGetProperty("input_groups", out var groups) && groups.ValueKind == JsonValueKind.Array)
                    {
                        foreach (var group in groups.EnumerateArray())
                        {
                            var detailGroup = new DetailGroup { Title = Text(group, "title") };
                            if (group.TryGetProperty("inputs", out var inputs) && inputs.ValueKind == JsonValueKind.Array)
                            {
                                foreach (var input in inputs.EnumerateArray())
                                {
                                    var field = ReadField(input, answers, uploads);
                                    if (field != null)
                                    {
                                        detailGroup.Fields.Add(field);
                                    }
                                }
                            }
                            if (detailGroup.Fields.Count > 0 || detailGroup.Title.Length > 0)
                            {
                                detailSection.Groups.Add(detailGroup);
                            }
                        }
                    }
                    if (detailSection.Groups.Count > 0)
                    {
                        details.Sections.Add(detailSection);
                    }
                }
                return details;
            }
        }

        private static DetailField? ReadField(JsonElement input, JsonElement? answers, List<DetailUpload> uploads)
        {
            var type = Text(input, "type");
            var id = Text(input, "id");
            var label = Text(input, "label");
            JsonElement answer = default;
            var hasAnswer = answers != null && id.Length > 0 && answers.Value.TryGetProperty(id, out answer) && answer.ValueKind != JsonValueKind.Null;

            switch (type)
            {
                case "markdown_text":
                case "text_end_link":
                    return label.Length == 0 ? null : new DetailField { Kind = DetailFieldKind.Note, Label = label };
                case "checkbox":
                    return new DetailField { Kind = DetailFieldKind.Tick, Label = label, Ticked = hasAnswer && answer.ValueKind == JsonValueKind.True };
                case "verifiable_checkbox":
                    return new DetailField
                    {
                        Kind = DetailFieldKind.Tick,
                        Label = label,
                        Ticked = hasAnswer && answer.ValueKind == JsonValueKind.True,
                        Verification = Verification(answers, id)
                    };
                case "file_uploader":
                    // The question's answer lists the ids of its files; the files themselves (name, date) are in the record's uploads
                    return new DetailField { Kind = DetailFieldKind.Files, Label = label, Files = hasAnswer ? FilesChosen(answer, uploads) : new List<DetailUpload>() };
                case "radio":
                case "single_tab_selector":
                case "multi_tab_selector":
                    return new DetailField { Kind = DetailFieldKind.Text, Label = label, Value = hasAnswer ? ChosenTitles(input, answer) : string.Empty };
                default:
                    return new DetailField { Kind = DetailFieldKind.Text, Label = label, Value = hasAnswer ? Display(answer) : string.Empty };
            }
        }

        private static List<DetailUpload> FilesChosen(JsonElement answer, List<DetailUpload> uploads)
        {
            if (answer.ValueKind != JsonValueKind.Array)
            {
                return new List<DetailUpload>();
            }
            var ids = new HashSet<string>(answer.EnumerateArray().Where(a => a.ValueKind == JsonValueKind.String).Select(a => a.GetString() ?? string.Empty));
            return uploads.Where(u => ids.Contains(u.Id)).ToList();
        }

        // Once a leader has signed a requirement off, the record holds who, how to reach them and when, beside the tick:
        // <id>_verifierName, <id>_verifierContact and <id>_verifiedDate. Terrain shows it as "Verified by X (contact). Date D".
        private static string Verification(JsonElement? answers, string id)
        {
            if (answers == null || id.Length == 0)
            {
                return string.Empty;
            }
            var name = Text(answers.Value, id + "_verifierName");
            if (name.Length == 0)
            {
                return string.Empty;
            }
            var contact = Text(answers.Value, id + "_verifierContact");
            var date = Text(answers.Value, id + "_verifiedDate");
            return $"Verified by {name}" + (contact.Length > 0 ? $" ({contact})" : "") + ". " + (date.Length > 0 ? $"Date {date}" : "").Trim();
        }

        // The titles of the chosen options (the answer holds their ids)
        private static string ChosenTitles(JsonElement input, JsonElement answer)
        {
            var titles = new Dictionary<string, string>();
            void Collect(JsonElement selections)
            {
                if (selections.ValueKind != JsonValueKind.Array)
                {
                    return;
                }
                foreach (var selection in selections.EnumerateArray())
                {
                    var id = Text(selection, "id");
                    if (id.Length > 0)
                    {
                        titles[id] = Text(selection, "title");
                    }
                }
            }
            if (input.TryGetProperty("selections", out var direct))
            {
                Collect(direct);
            }
            if (input.TryGetProperty("tabs", out var tabs) && tabs.ValueKind == JsonValueKind.Array)
            {
                foreach (var tab in tabs.EnumerateArray())
                    if (tab.TryGetProperty("selections", out var tabSelections))
                    {
                        Collect(tabSelections);
                    }
            }

            var chosen = answer.ValueKind == JsonValueKind.Array
                ? answer.EnumerateArray().Select(Display)
                : new[] { Display(answer) };
            return string.Join(", ", chosen.Where(c => c.Length > 0).Select(c => titles.TryGetValue(c, out var title) && title.Length > 0 ? title : c));
        }

        private static string Display(JsonElement value)
        {
            switch (value.ValueKind)
            {
                case JsonValueKind.String: return value.GetString() ?? string.Empty;
                case JsonValueKind.Number: return value.GetRawText();
                case JsonValueKind.True: return "Yes";
                case JsonValueKind.False: return "No";
                case JsonValueKind.Null:
                case JsonValueKind.Undefined: return string.Empty;
                case JsonValueKind.Array: return string.Join(", ", value.EnumerateArray().Select(Display).Where(v => v.Length > 0));
                default: return string.Join(", ", value.EnumerateObject().Select(p => $"{p.Name}: {Display(p.Value)}"));
            }
        }

        private static List<DetailUpload> ReadUploads(JsonElement achievement)
        {
            var uploads = new List<DetailUpload>();
            if (achievement.TryGetProperty("uploads", out var list) && list.ValueKind == JsonValueKind.Array)
            {
                foreach (var upload in list.EnumerateArray())
                {
                    uploads.Add(new DetailUpload
                    {
                        Id = Text(upload, "id"),
                        FileName = Text(upload, "filename"),
                        UploadedOn = DateTime.TryParse(Text(upload, "uploaded_on"), null, System.Globalization.DateTimeStyles.RoundtripKind, out var parsed) ? parsed : null
                    });
                }
            }
            return uploads;
        }

        private static string Text(JsonElement element, string name) =>
            element.ValueKind == JsonValueKind.Object && element.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.String
                ? value.GetString() ?? string.Empty
                : string.Empty;

        private static bool TryParse(string? json, out JsonDocument? document)
        {
            document = null;
            if (string.IsNullOrWhiteSpace(json))
            {
                return false;
            }
            try
            {
                document = JsonDocument.Parse(json);
                return true;
            }
            catch (JsonException)
            {
                return false;
            }
        }
    }

    public class DetailSection
    {
        public string Title { get; init; } = string.Empty;
        public List<DetailGroup> Groups { get; init; } = new();
    }

    public class DetailGroup
    {
        public string Title { get; init; } = string.Empty;
        public List<DetailField> Fields { get; init; } = new();
    }

    public enum DetailFieldKind
    {
        Text,
        Tick,
        Files,
        Note
    }

    public class DetailField
    {
        public DetailFieldKind Kind { get; init; }
        public string Label { get; init; } = string.Empty;
        public string Value { get; init; } = string.Empty;
        public bool Ticked { get; init; }
        public string Verification { get; init; } = string.Empty;
        public List<DetailUpload> Files { get; init; } = new();
    }

    public class DetailUpload
    {
        public string Id { get; init; } = string.Empty;
        public string FileName { get; init; } = string.Empty;
        public DateTime? UploadedOn { get; init; }
    }
}
