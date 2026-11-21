using System.Text.Json;
using Topo.Model.BadgeStore;
using Topo.Model.Program;

namespace TopoReportFunctionTest
{
    [TestClass]
    public class BadgeStoreWritesTests
    {
        private const string Group = "group-1";

        private static JsonElement Body() =>
            JsonDocument.Parse(JsonSerializer.Serialize(BadgeStoreWrites.CreateEventBody(Group, "me", "Pat Leader"))).RootElement;

        [TestMethod]
        public void TheStoreEventIsAPlannedGroupEventOnTheFixedDayWithNobodyAttending()
        {
            var body = Body();
            Assert.AreEqual(BadgeStoreSpec.Title, body.GetProperty("title").GetString());
            Assert.AreEqual(BadgeStoreSpec.Description, body.GetProperty("description").GetString());
            Assert.AreEqual("planned", body.GetProperty("status").GetString());
            Assert.AreEqual("group", body.GetProperty("event_type").GetProperty("type").GetString());
            Assert.AreEqual(Group, body.GetProperty("event_type").GetProperty("id").GetString());
            Assert.AreEqual("2100-02-01T08:00:00.000+00:00", body.GetProperty("start_datetime").GetString());
            Assert.AreEqual("2100-02-01T09:00:00.000+00:00", body.GetProperty("end_datetime").GetString());
            var attendance = body.GetProperty("attendance");
            foreach (var list in new[] { "leader_member_ids", "assistant_member_ids", "attendee_member_ids", "participant_member_ids" })
                Assert.AreEqual(0, attendance.GetProperty(list).GetArrayLength(), list);
        }

        [TestMethod]
        public void TheOrganiserIsTheLeaderCreatingItAndIsTheVerifier()
        {
            var body = Body();
            Assert.AreEqual("me", body.GetProperty("organisers")[0].GetString());
            Assert.AreEqual("Pat Leader", body.GetProperty("achievement_pathway_oas_data").GetProperty("verifier").GetProperty("name").GetString());
            Assert.AreEqual("Pat Leader", body.GetProperty("achievement_pathway_logbook_data").GetProperty("verifier").GetProperty("name").GetString());
        }

        [TestMethod]
        public void TheBodyHasEveryFieldTerrainNeededForAPlannedEvent()
        {
            var body = Body();
            foreach (var name in new[] { "title", "description", "justification", "organisers", "challenge_area", "start_datetime", "end_datetime", "event_type",
                         "attendance", "schedule_items", "achievement_pathway_oas_data", "achievement_pathway_logbook_data", "review", "uploads",
                         "equipment_notes", "additional_notes", "location", "iana_timezone", "status" })
                Assert.IsTrue(body.TryGetProperty(name, out _), name);
        }

        private static Event Store(string id, string group = Group, string title = BadgeStoreSpec.Title, string start = "2100-02-01T08:00:00Z", string status = "planned") =>
            new Event { id = id, group_id = group, title = title, start_datetime = DateTime.Parse(start).ToUniversalTime(), status = status };

        [TestMethod]
        public void WithSeveralStoreEventsTheSameOneIsAlwaysWrittenTo()
        {
            var a = Store("a111");
            var b = Store("b222");
            Assert.AreEqual("a111", BadgeStoreWrites.ChooseWriteEvent(new[] { b, a })!.id);
            Assert.AreEqual("a111", BadgeStoreWrites.ChooseWriteEvent(new[] { a, b })!.id);
            Assert.IsNull(BadgeStoreWrites.ChooseWriteEvent(Array.Empty<Event>()));
        }

        [TestMethod]
        public void TheEventJustCreatedIsRecognisedInTheOrganisedList()
        {
            Assert.IsTrue(BadgeStoreWrites.IsCreatedStoreEvent(Store("e1"), Group));
            Assert.IsFalse(BadgeStoreWrites.IsCreatedStoreEvent(Store("e1", group: "other"), Group));
            Assert.IsFalse(BadgeStoreWrites.IsCreatedStoreEvent(Store("e1", title: "Camp"), Group));
            Assert.IsFalse(BadgeStoreWrites.IsCreatedStoreEvent(Store("e1", start: "2026-02-01T08:00:00Z"), Group));
            Assert.IsFalse(BadgeStoreWrites.IsCreatedStoreEvent(Store("e1", status: "cancelled"), Group));
        }

        private static BadgeStoreContents Contents(string member, string award, string awarded, string? presented) => new BadgeStoreContents
        {
            Records = new List<BadgeStoreRecord>
            {
                new BadgeStoreRecord { MemberId = member, AwardId = award, Awarded = DateTime.Parse(awarded), Presented = presented == null ? null : DateTime.Parse(presented) }
            }
        };

        [TestMethod]
        public void RepeatingWhatTheStoreAlreadySaysIsNotPostedAgain()
        {
            var contents = Contents("m1", "camper_award_10", "2026-08-26", "2026-09-02");
            Assert.IsTrue(BadgeStoreWrites.IsAlreadyRecorded(contents, "m1", "camper_award_10", new DateTime(2026, 8, 26), new DateTime(2026, 9, 2)));
        }

        [TestMethod]
        public void ADifferentDateMemberOrAwardIsPosted()
        {
            var contents = Contents("m1", "camper_award_10", "2026-08-26", "2026-09-02");
            Assert.IsFalse(BadgeStoreWrites.IsAlreadyRecorded(contents, "m1", "camper_award_10", new DateTime(2026, 8, 27), new DateTime(2026, 9, 2)));
            Assert.IsFalse(BadgeStoreWrites.IsAlreadyRecorded(contents, "m1", "camper_award_10", new DateTime(2026, 8, 26), new DateTime(2026, 9, 9)));
            Assert.IsFalse(BadgeStoreWrites.IsAlreadyRecorded(contents, "m2", "camper_award_10", new DateTime(2026, 8, 26), new DateTime(2026, 9, 2)));
            Assert.IsFalse(BadgeStoreWrites.IsAlreadyRecorded(contents, "m1", "camper_award_25", new DateTime(2026, 8, 26), new DateTime(2026, 9, 2)));
        }

        [TestMethod]
        public void AddingAPresentedDateToARecordWithoutOneIsPosted()
        {
            var contents = Contents("m1", "camper_award_10", "2026-08-26", null);
            Assert.IsFalse(BadgeStoreWrites.IsAlreadyRecorded(contents, "m1", "camper_award_10", new DateTime(2026, 8, 26), new DateTime(2026, 9, 2)));
        }

        [TestMethod]
        public void ARequestWithNoPresentedDateDoesNotTakeAwayOneAlreadyRecorded()
        {
            var contents = Contents("m1", "camper_award_10", "2026-08-26", "2026-09-02");
            Assert.IsTrue(BadgeStoreWrites.IsAlreadyRecorded(contents, "m1", "camper_award_10", new DateTime(2026, 8, 26), null));
        }

        [TestMethod]
        public void ARecordIsFoundOnlyForItsMemberAndAward()
        {
            var contents = Contents("m1", "camper_award_10", "2026-08-26", null);
            Assert.IsTrue(BadgeStoreWrites.HasRecord(contents, "m1", "camper_award_10"));
            Assert.IsFalse(BadgeStoreWrites.HasRecord(contents, "m1", "camper_award_25"));
            Assert.IsFalse(BadgeStoreWrites.HasRecord(contents, "m2", "camper_award_10"));
        }

        [TestMethod]
        public void ARecordKnowsWhichEventAndCommentItCameFrom()
        {
            var comment = new EventComment
            {
                id = "c1", EventId = "e1", status = "active", created_datetime = DateTime.UtcNow,
                text = BadgeStoreReader.AwardText("m1", "walkabout_10", new DateTime(2026, 8, 26), null), member = new EventCommentMember { id = "leader-1" }
            };
            var record = BadgeStoreReader.Resolve(new[] { comment }, _ => true).Records.Single();
            Assert.AreEqual("e1", record.EventId);
            Assert.AreEqual("c1", record.CommentId);
        }

        [TestMethod]
        public void OnlyAFailureOrAMissingStoreIsNotASuccess()
        {
            Assert.IsTrue(new BadgeStoreWriteResult { Outcome = BadgeStoreWriteOutcome.Written }.Succeeded);
            Assert.IsTrue(new BadgeStoreWriteResult { Outcome = BadgeStoreWriteOutcome.AlreadyRecorded }.Succeeded);
            Assert.IsTrue(new BadgeStoreWriteResult { Outcome = BadgeStoreWriteOutcome.StoreCreated }.Succeeded);
            Assert.IsFalse(new BadgeStoreWriteResult { Outcome = BadgeStoreWriteOutcome.Failed }.Succeeded);
            Assert.IsFalse(new BadgeStoreWriteResult { Outcome = BadgeStoreWriteOutcome.NoStoreEvent }.Succeeded);
        }
    }
}
