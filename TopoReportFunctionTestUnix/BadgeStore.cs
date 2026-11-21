using Topo.Model.BadgeStore;
using Topo.Model.Program;

namespace TopoReportFunctionTest
{
    [TestClass]
    public class BadgeStore
    {
        private const string Group = "group-1";
        private static readonly DateTime T0 = new DateTime(2026, 11, 1, 9, 0, 0, DateTimeKind.Utc);

        private static EventComment Comment(string id, string text, int minutes, string author = "leader-1", string status = "active") => new EventComment
        {
            id = id, text = text, created_datetime = T0.AddMinutes(minutes), status = status, member = new EventCommentMember { id = author }
        };

        private static string Award(string member, string award, string awarded = "2026-08-26", string? presented = null) =>
            BadgeStoreReader.AwardText(member, award, DateTime.Parse(awarded), presented == null ? null : DateTime.Parse(presented));

        private static BadgeStoreContents Read(params EventComment[] comments) => BadgeStoreReader.Resolve(comments, a => a.StartsWith("leader"));

        [TestMethod]
        public void AnAwardCommentBecomesARecord()
        {
            var contents = Read(Comment("c1", Award("m1", "camper_award_10", "2026-08-26", "2026-09-02"), 0));
            var record = contents.Records.Single();
            Assert.AreEqual("m1", record.MemberId);
            Assert.AreEqual("camper_award_10", record.AwardId);
            Assert.AreEqual(new DateTime(2026, 8, 26), record.Awarded);
            Assert.AreEqual(new DateTime(2026, 9, 2), record.Presented);
            Assert.AreEqual("leader-1", record.RecordedById);
        }

        [TestMethod]
        public void ThePresentedDateIsOptional()
        {
            var record = Read(Comment("c1", Award("m1", "walkabout_10"), 0)).Records.Single();
            Assert.IsNull(record.Presented);
        }

        [TestMethod]
        public void TheLatestCommentForAMemberAndAwardWins()
        {
            var contents = Read(
                Comment("c1", Award("m1", "walkabout_10", "2026-08-26"), 0),
                Comment("c2", Award("m1", "walkabout_10", "2026-08-26", "2026-09-02"), 5));
            Assert.AreEqual(1, contents.Records.Count);
            Assert.AreEqual(new DateTime(2026, 9, 2), contents.Records[0].Presented);
        }

        [TestMethod]
        public void TheResultDoesNotDependOnTheOrderTerrainReturnsComments()
        {
            var first = Comment("c1", Award("m1", "walkabout_10", "2026-08-26"), 0);
            var second = Comment("c2", Award("m1", "walkabout_10", "2026-09-30"), 5);
            Assert.AreEqual(new DateTime(2026, 9, 30), Read(second, first).Records.Single().Awarded);
            Assert.AreEqual(new DateTime(2026, 9, 30), Read(first, second).Records.Single().Awarded);
        }

        [TestMethod]
        public void ARevokeRemovesTheRecordAndALaterAwardBringsItBack()
        {
            var award = Comment("c1", Award("m1", "camper_award_10"), 0);
            var revoke = Comment("c2", BadgeStoreReader.RevokeText("m1", "camper_award_10"), 5);
            Assert.AreEqual(0, Read(award, revoke).Records.Count);
            var again = Comment("c3", Award("m1", "camper_award_10", "2026-10-01"), 10);
            Assert.AreEqual(new DateTime(2026, 10, 1), Read(award, revoke, again).Records.Single().Awarded);
        }

        [TestMethod]
        public void ARevokeOnlyAffectsItsOwnMemberAndAward()
        {
            var contents = Read(
                Comment("c1", Award("m1", "camper_award_10"), 0),
                Comment("c2", Award("m2", "camper_award_10"), 1),
                Comment("c3", Award("m1", "camper_award_25"), 2),
                Comment("c4", BadgeStoreReader.RevokeText("m1", "camper_award_10"), 3));
            CollectionAssert.AreEqual(new[] { "m1/camper_award_25", "m2/camper_award_10" }, contents.Records.Select(r => r.MemberId + "/" + r.AwardId).ToArray());
        }

        [TestMethod]
        public void CommentsFromSomeoneNotTrustedAreIgnoredAndCounted()
        {
            var contents = Read(
                Comment("c1", Award("m1", "camper_award_10"), 0, author: "parent-9"),
                Comment("c2", Award("m2", "camper_award_10"), 1));
            Assert.AreEqual("m2", contents.Records.Single().MemberId);
            Assert.AreEqual(1, contents.IgnoredUntrustedAuthor);
        }

        [TestMethod]
        public void AnUntrustedRevokeCannotRemoveARecord()
        {
            var contents = Read(
                Comment("c1", Award("m1", "camper_award_10"), 0),
                Comment("c2", BadgeStoreReader.RevokeText("m1", "camper_award_10"), 1, author: "parent-9"));
            Assert.AreEqual(1, contents.Records.Count);
        }

        [TestMethod]
        public void CommentsThatAreNotRecordsAreCountedNotUsed()
        {
            var contents = Read(
                Comment("c1", "hello", 0),
                Comment("c2", "[1,2]", 1),
                Comment("c3", "{\"v\":1,\"t\":\"award\",\"m\":\"m1\",\"a\":\"camper_award_10\",\"d\":\"not a date\"}", 2),
                Comment("c4", "{\"v\":1,\"t\":\"award\",\"a\":\"camper_award_10\",\"d\":\"2026-08-26\"}", 3),
                Comment("c5", "{\"v\":1,\"t\":\"mystery\",\"m\":\"m1\",\"a\":\"x\"}", 4),
                Comment("c6", "{\"v\":1,\"t\":\"award\",\"m\":\"m1\",\"a\":\"camper_award_10\",\"d\":\"2026-08-26\",\"p\":\"soon\"}", 5));
            Assert.AreEqual(0, contents.Records.Count);
            Assert.AreEqual(6, contents.IgnoredUnreadable);
        }

        [TestMethod]
        public void ARecordOfANewerVersionIsLeftAloneAndCounted()
        {
            var contents = Read(Comment("c1", "{\"v\":2,\"t\":\"award\",\"m\":\"m1\",\"a\":\"camper_award_10\",\"d\":\"2026-08-26\"}", 0));
            Assert.AreEqual(0, contents.Records.Count);
            Assert.AreEqual(1, contents.IgnoredNewerVersion);
        }

        [TestMethod]
        public void DeletedCommentsAreIgnored()
        {
            var contents = Read(Comment("c1", Award("m1", "camper_award_10"), 0, status: "deleted"));
            Assert.AreEqual(0, contents.Records.Count);
            Assert.AreEqual(0, contents.IgnoredUnreadable);
        }

        [TestMethod]
        public void TheTextWrittenIsTheTextRead()
        {
            var text = BadgeStoreReader.AwardText("a\"b", "walkabout_10", new DateTime(2026, 8, 26), null);
            Assert.AreEqual("{\"v\":1,\"t\":\"award\",\"m\":\"a\\u0022b\",\"a\":\"walkabout_10\",\"d\":\"2026-08-26\"}", text);
            Assert.AreEqual("a\"b", Read(Comment("c1", text, 0)).Records.Single().MemberId);
            Assert.AreEqual("{\"v\":1,\"t\":\"revoke\",\"m\":\"m1\",\"a\":\"walkabout_10\"}", BadgeStoreReader.RevokeText("m1", "walkabout_10"));
        }

        private static Event StoreEvent(string title = BadgeStoreSpec.Title, string type = "group", string invitee = Group, string status = "planned") =>
            new Event { id = "e1", title = title, invitee_type = type, invitee_id = invitee, status = status };

        [TestMethod]
        public void TheStoreEventIsTheGroupsOwnEventWithTheStoreTitle()
        {
            Assert.IsTrue(BadgeStoreSpec.IsStoreEvent(StoreEvent(), Group));
            Assert.IsTrue(BadgeStoreSpec.IsStoreEvent(StoreEvent(title: "topo RECORDS (renamed)"), Group));
            Assert.IsFalse(BadgeStoreSpec.IsStoreEvent(StoreEvent(invitee: "another-group"), Group));
            Assert.IsFalse(BadgeStoreSpec.IsStoreEvent(StoreEvent(type: "unit"), Group));
            Assert.IsFalse(BadgeStoreSpec.IsStoreEvent(StoreEvent(title: "Summit Award Storage - Please Ignore"), Group));
            Assert.IsFalse(BadgeStoreSpec.IsStoreEvent(StoreEvent(status: "cancelled"), Group));
        }

        [TestMethod]
        public void TheStoreEventSitsInsideItsSearchWindow()
        {
            Assert.IsTrue(BadgeStoreSpec.StartUtc > BadgeStoreSpec.SearchFromUtc && BadgeStoreSpec.EndUtc < BadgeStoreSpec.SearchToUtc);
            Assert.IsTrue(BadgeStoreSpec.Description.StartsWith("Topo records v" + BadgeStoreSpec.SchemaVersion));
        }

        private static GetCalendarsResultModel Calendars(bool groupOn) => new GetCalendarsResultModel
        {
            member_id = "me",
            own_calendars = new[]
            {
                new Own_Calendars { id = Group, type = "group", title = "The group", selected = groupOn },
                new Own_Calendars { id = "unit-1", type = "unit", title = "Scout Unit", selected = true, section = "scout" },
                new Own_Calendars { id = "patrol-1", type = "patrol", title = "Adult Leaders", selected = false }
            },
            other_calendars = new[] { new Other_Calendars { id = "other-1", type = "unit", title = "Another unit", selected = false } }
        };

        [TestMethod]
        public void SwitchingTheGroupCalendarOnChangesOnlyThatAndLeavesTheOriginalAlone()
        {
            var original = Calendars(false);
            var changed = BadgeStoreCalendars.WithGroupCalendarOn(original, Group)!;
            CollectionAssert.AreEqual(new[] { true, true, false }, changed.own_calendars.Select(c => c.selected).ToArray());
            CollectionAssert.AreEqual(new[] { false, true, false }, original.own_calendars.Select(c => c.selected).ToArray());
            Assert.AreEqual("The group", changed.own_calendars[0].title);
            Assert.AreEqual("scout", changed.own_calendars[1].section);
            Assert.IsFalse(changed.other_calendars.Single().selected);
        }

        [TestMethod]
        public void NothingIsSwitchedWhenTheGroupCalendarIsAlreadyOnOrMissing()
        {
            Assert.IsNull(BadgeStoreCalendars.WithGroupCalendarOn(Calendars(true), Group));
            Assert.IsNull(BadgeStoreCalendars.WithGroupCalendarOn(Calendars(false), "another-group"));
        }
    }
}
