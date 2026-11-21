using Topo.Model.Approvals;

namespace TopoReportFunctionTest
{
    [TestClass]
    public class PendingApprovalsTests
    {
        private static PendingApprovalItem Item(string id, string type = "review") => new()
        {
            SubmissionId = id,
            SubmissionType = type,
            MemberName = "Member " + id,
            AchievementName = "Milestone 1"
        };

        private static PendingApprovalsSession Session(params string[] ids) => new(ids.Select(i => Item(i)));

        private static string? Ids(PendingApprovalsSession session) => session.Current?.SubmissionId;

        private static Task<string?> Works(PendingApprovalItem item, ApprovalAction action) => Task.FromResult<string?>(null);

        private static Task<string?> Fails(PendingApprovalItem item, ApprovalAction action) => Task.FromResult<string?>("Terrain said no.");

        // ---- order, position and progress ----

        [TestMethod]
        public void ItStartsOnTheFirstItemWithProgressOfOneOfTheTotal()
        {
            var session = Session("a", "b", "c");
            Assert.AreEqual("a", Ids(session));
            Assert.AreEqual(1, session.Position);
            Assert.AreEqual(3, session.Total);
            Assert.IsFalse(session.CanGoBack);
        }

        [TestMethod]
        public async Task ActioningItemsVisitsThemInOrderAndProgressFollows()
        {
            var session = Session("a", "b", "c");
            Assert.IsTrue(await session.CompleteAsync(ApprovalAction.Approve(""), Works));
            Assert.AreEqual("b", Ids(session));
            Assert.AreEqual(2, session.Position);
            Assert.AreEqual(3, session.Total);
            Assert.IsTrue(await session.CompleteAsync(ApprovalAction.RequestImprovement("redo"), Works));
            Assert.AreEqual("c", Ids(session));
            Assert.AreEqual(3, session.Position);
        }

        [TestMethod]
        public void TheQueueIsOldestFirstThenByName()
        {
            var pending = new[]
            {
                Pending("s3", "Zed", "Zhang", new DateTime(2026, 11, 3)),
                Pending("s1", "Amy", "Brown", new DateTime(2026, 11, 1)),
                Pending("s2b", "Ben", "Young", new DateTime(2026, 11, 2)),
                Pending("s2a", "Alf", "Young", new DateTime(2026, 11, 2)),
            };
            var queue = PendingApprovalQueue.FromPending(pending, "leader");
            CollectionAssert.AreEqual(new[] { "s1", "s2a", "s2b", "s3" }, queue.Items.Select(i => i.SubmissionId).ToArray());
        }

        // ---- skip ----

        [TestMethod]
        public void SkippingMovesOnAndLeavesTheItemPending()
        {
            var session = Session("a", "b", "c");
            session.Skip();
            Assert.AreEqual("b", Ids(session));
            Assert.AreEqual(2, session.Position);
            Assert.AreEqual(3, session.RemainingCount);
            Assert.AreEqual(1, session.SkippedCount);
            Assert.AreEqual(0, session.Completed.Count);
        }

        [TestMethod]
        public async Task SkippedItemsAreStepsOverWhenMovingOnAfterAnAction()
        {
            var session = Session("a", "b", "c", "d");
            session.Skip();                      // a skipped, now on b
            await session.CompleteAsync(ApprovalAction.Approve(""), Works); // b done, on c
            session.Skip();                      // c skipped, now on d
            await session.CompleteAsync(ApprovalAction.Approve(""), Works); // d done, end
            Assert.IsNull(session.Current);
            Assert.IsTrue(session.IsAtEnd);
            CollectionAssert.AreEqual(new[] { "a", "c" }, session.Remaining.Select(i => i.SubmissionId).ToArray());
        }

        // ---- back ----

        [TestMethod]
        public void BackGoesToTheEarlierItemAndForwardAgainIsSkip()
        {
            var session = Session("a", "b", "c");
            session.Skip();
            session.Skip();
            Assert.AreEqual("c", Ids(session));
            session.Back();
            Assert.AreEqual("b", Ids(session));
            session.Back();
            Assert.AreEqual("a", Ids(session));
            Assert.IsFalse(session.CanGoBack);
            session.Back();
            Assert.AreEqual("a", Ids(session), "Back on the first item stays put");
        }

        [TestMethod]
        public async Task BackNeverReturnsToAnItemThatWasActioned()
        {
            var session = Session("a", "b", "c");
            await session.CompleteAsync(ApprovalAction.Approve(""), Works); // a done
            session.Skip();                                                  // b skipped, on c
            session.Back();
            Assert.AreEqual("b", Ids(session));
            Assert.IsFalse(session.CanGoBack, "a is done, so there is nothing before b");
        }

        [TestMethod]
        public void BackFromTheEndReturnsToTheLastItemStillWaiting()
        {
            var session = Session("a", "b");
            session.Skip();
            session.Skip();
            Assert.IsTrue(session.IsAtEnd);
            Assert.IsTrue(session.CanGoBack);
            session.Back();
            Assert.AreEqual("b", Ids(session));
        }

        // ---- the end ----

        [TestMethod]
        public async Task ActioningEverythingEndsAllDone()
        {
            var session = Session("a", "b");
            await session.CompleteAsync(ApprovalAction.Approve(""), Works);
            await session.CompleteAsync(ApprovalAction.RequestImprovement(""), Works);
            Assert.IsNull(session.Current);
            Assert.IsTrue(session.IsAtEnd);
            Assert.IsTrue(session.IsAllDone);
            Assert.AreEqual(2, session.Completed.Count);
            Assert.IsFalse(session.CanGoBack);
        }

        [TestMethod]
        public void AnEmptyQueueIsAlreadyAtTheEnd()
        {
            var session = Session();
            Assert.IsNull(session.Current);
            Assert.IsTrue(session.IsAtEnd);
            Assert.IsTrue(session.IsAllDone);
            Assert.AreEqual(0, session.Total);
        }

        [TestMethod]
        public void SkippingTheLastItemEndsWithTheSkippedOnesStillPendingAndRestartReturnsToThem()
        {
            var session = Session("a", "b");
            session.Skip();
            session.Skip();
            Assert.IsTrue(session.IsAtEnd);
            Assert.IsFalse(session.IsAllDone);
            Assert.AreEqual(2, session.RemainingCount);
            session.Restart();
            Assert.AreEqual("a", Ids(session));
        }

        [TestMethod]
        public async Task CompleteAndSkipDoNothingOnceThereIsNoCurrentItem()
        {
            var session = Session();
            var called = false;
            Assert.IsFalse(await session.CompleteAsync(ApprovalAction.Approve(""), (i, a) => { called = true; return Task.FromResult<string?>(null); }));
            Assert.IsFalse(called);
            session.Skip();
            session.Back();
            Assert.IsNull(session.Current);
        }

        // ---- stop ----

        [TestMethod]
        public async Task StoppingKeepsWhatWasDoneAndShowsNoCurrentItem()
        {
            var session = Session("a", "b", "c");
            await session.CompleteAsync(ApprovalAction.Approve(""), Works);
            session.Stop();
            Assert.IsTrue(session.IsStopped);
            Assert.IsNull(session.Current);
            Assert.IsFalse(session.IsAtEnd, "stopped is not the same as reaching the end");
            Assert.AreEqual(1, session.Completed.Count);
            Assert.AreEqual(2, session.RemainingCount);
            Assert.IsFalse(session.CanGoBack);
        }

        [TestMethod]
        public async Task NothingIsSentOnceStopped()
        {
            var session = Session("a", "b");
            session.Stop();
            var called = false;
            Assert.IsFalse(await session.CompleteAsync(ApprovalAction.Approve(""), (i, a) => { called = true; return Task.FromResult<string?>(null); }));
            Assert.IsFalse(called);
            session.Skip();
            session.Back();
            session.Restart();
            Assert.IsNull(session.Current);
        }

        // ---- a failed action ----

        [TestMethod]
        public async Task AFailedActionDoesNotAdvanceAndSaysWhy()
        {
            var session = Session("a", "b");
            Assert.IsFalse(await session.CompleteAsync(ApprovalAction.Approve("ok"), Fails));
            Assert.AreEqual("a", Ids(session));
            Assert.AreEqual(1, session.Position);
            Assert.AreEqual(2, session.RemainingCount);
            Assert.AreEqual(0, session.Completed.Count);
            Assert.AreEqual("Terrain said no.", session.LastError);
        }

        [TestMethod]
        public async Task AThrownErrorCountsAsAFailureToo()
        {
            var session = Session("a", "b");
            var worked = await session.CompleteAsync(ApprovalAction.Approve(""), (i, a) => throw new InvalidOperationException("boom"));
            Assert.IsFalse(worked);
            Assert.AreEqual("a", Ids(session));
            Assert.AreEqual(0, session.Completed.Count);
            Assert.AreEqual("boom", session.LastError);
        }

        [TestMethod]
        public async Task ARetryAfterAFailureCanWorkAndClearsTheError()
        {
            var session = Session("a", "b");
            await session.CompleteAsync(ApprovalAction.Approve(""), Fails);
            Assert.IsTrue(await session.CompleteAsync(ApprovalAction.Approve(""), Works));
            Assert.AreEqual("b", Ids(session));
            Assert.IsNull(session.LastError);
        }

        [TestMethod]
        public async Task TheActionIsDoneToTheCurrentItemAndEachItemIsSentOnce()
        {
            var session = Session("a", "b", "c");
            session.Skip();
            var sent = new List<string>();
            await session.CompleteAsync(ApprovalAction.RequestImprovement("more"), (i, a) => { sent.Add($"{i.SubmissionId}:{a.Outcome}"); return Task.FromResult<string?>(null); });
            CollectionAssert.AreEqual(new[] { "b:rejected" }, sent);
        }

        // ---- items dealt with elsewhere ----

        [TestMethod]
        public void ItemsNoLongerPendingInTerrainAreDroppedAndProgressShrinks()
        {
            var session = Session("a", "b", "c", "d");
            var dropped = session.Reconcile(new[] { "a", "d" });
            Assert.AreEqual(2, dropped);
            Assert.AreEqual(2, session.Total);
            Assert.AreEqual("a", Ids(session));
            session.Skip();
            Assert.AreEqual("d", Ids(session));
            Assert.AreEqual(2, session.Position);
        }

        [TestMethod]
        public void IfTheItemOnScreenWasDealtWithElsewhereTheListMovesToTheNext()
        {
            var session = Session("a", "b");
            session.Reconcile(new[] { "b" });
            Assert.AreEqual("b", Ids(session));
            Assert.AreEqual(1, session.Position);
            Assert.AreEqual(1, session.Total);
        }

        // ---- what Terrain is sent ----

        [TestMethod]
        public void ApproveAndImproveSendTheOutcomeAndTheTrimmedComment()
        {
            var approve = ApprovalAction.Approve("  well done  ").ToRequestBody();
            Assert.AreEqual("approved", approve["outcome"]);
            Assert.AreEqual("well done", approve["comment"]);

            var improve = ApprovalAction.RequestImprovement(null).ToRequestBody();
            Assert.AreEqual("rejected", improve["outcome"]);
            Assert.AreEqual("", improve["comment"]);
        }

        [TestMethod]
        public void ACommentOverFiveHundredCharactersIsRefused()
        {
            Assert.IsNull(ApprovalAction.Approve(new string('x', 500)).Problem("review"));
            Assert.IsNotNull(ApprovalAction.Approve(new string('x', 501)).Problem("review"));
            Assert.IsNotNull(ApprovalAction.RequestImprovement(new string('x', 501)).Problem("approval"));
        }

        [TestMethod]
        public void ApproveAndImproveAreOfferedForReviewAndApprovalRequestsButNeverForAwards()
        {
            Assert.IsNull(ApprovalAction.Approve("").Problem("review"));
            Assert.IsNull(ApprovalAction.Approve("").Problem("approval"));
            Assert.IsNull(ApprovalAction.RequestImprovement("").Problem("approval"));
            Assert.IsNotNull(ApprovalAction.Approve("").Problem("award"));
            Assert.IsNotNull(ApprovalAction.RequestImprovement("").Problem("award"));
        }

        // ---- which pending requests are offered ----

        private static Result Pending(string submissionId, string first, string last, DateTime date, string memberId = "m", string type = "review", params string[] actionedBy)
        {
            return new Result
            {
                achievement = new Achievement { id = "ach-" + submissionId, type = "milestone", achievement_meta = new Achievement_Meta { stage = 2 } },
                submission = new Submission
                {
                    id = submissionId,
                    type = type,
                    date = date,
                    status = "pending",
                    actioned_by = actionedBy.Select(m => new Actioned_By { member_id = m, outcome = "approved", time = date }).ToArray()
                },
                member = new Member { id = memberId + "-" + submissionId, first_name = first, last_name = last }
            };
        }

        [TestMethod]
        public void OwnRequestsAndOnesAlreadyActionedAreLeftOutAndCounted()
        {
            var own = Pending("s1", "Lee", "Leader", new DateTime(2026, 11, 1));
            own.member.id = "leader";
            var alreadyDone = Pending("s2", "Amy", "Brown", new DateTime(2026, 11, 2), actionedBy: "leader");
            var offered = Pending("s3", "Ben", "Green", new DateTime(2026, 11, 3), actionedBy: "someone-else");

            var queue = PendingApprovalQueue.FromPending(new[] { own, alreadyDone, offered }, "leader");

            CollectionAssert.AreEqual(new[] { "s3" }, queue.Items.Select(i => i.SubmissionId).ToArray());
            Assert.AreEqual(2, queue.NotActionable);
        }

        [TestMethod]
        public void AwardRequestsAreLeftOutOfPendingApprovals()
        {
            var review = Pending("s1", "Amy", "Brown", new DateTime(2026, 11, 1));
            var approval = Pending("s2", "Ben", "Green", new DateTime(2026, 11, 2), type: "approval");
            var award = Pending("s3", "Cal", "White", new DateTime(2026, 11, 3), type: "award");

            var queue = PendingApprovalQueue.FromPending(new[] { review, approval, award }, "leader");

            CollectionAssert.AreEqual(new[] { "s1", "s2" }, queue.Items.Select(i => i.SubmissionId).ToArray());
            Assert.AreEqual(0, queue.NotActionable);
        }

        [TestMethod]
        public void WhenTheLeadersOwnIdIsUnknownNothingIsOffered()
        {
            var queue = PendingApprovalQueue.FromPending(new[] { Pending("s1", "Amy", "Brown", new DateTime(2026, 11, 1)) }, null);
            Assert.AreEqual(0, queue.Items.Count);
            Assert.AreEqual(1, queue.NotActionable);
        }

        [TestMethod]
        public void AnItemCarriesWhatTheLeaderNeedsToDecide()
        {
            var result = Pending("s1", "Amy", "Brown", new DateTime(2026, 11, 1), actionedBy: "other");
            result.submitted_by = new Submitted_By { id = "x", first_name = "Sam", last_name = "Smith" };
            result.submission.actioned_by[0].member_first_name = "Pat";
            result.submission.actioned_by[0].member_last_name = "Jones";
            result.submission.actioned_by[0].comment = "Looks good";

            var item = PendingApprovalQueue.FromPending(new[] { result }, "leader").Items.Single();

            Assert.AreEqual("Amy Brown", item.MemberName);
            Assert.AreEqual("Milestone 2", item.AchievementName);
            Assert.AreEqual("Sam Smith", item.SubmittedBy);
            Assert.AreEqual("Pat Jones", item.PreviousActions.Single().By);
            Assert.AreEqual("Looks good", item.PreviousActions.Single().Comment);
        }

        [TestMethod]
        public void AchievementNamesMatchTheApprovalsList()
        {
            Assert.AreEqual("Milestone 3", ApprovalNames.AchievementName("milestone", new Achievement_Meta { stage = 3 }));
            Assert.AreEqual("OAS Bushwalking 2", ApprovalNames.AchievementName("outdoor_adventure_skill", new Achievement_Meta { stream = "bushwalking", branch = "bushwalking", stage = 2 }));
            Assert.AreEqual("OAS Aquatics Boating 4", ApprovalNames.AchievementName("outdoor_adventure_skill", new Achievement_Meta { stream = "aquatics", branch = "boating", stage = 4 }));
            Assert.AreEqual("Intro To Scouting", ApprovalNames.AchievementName("intro_scouting", null));
            Assert.AreEqual("Special Interest Area", ApprovalNames.AchievementName("special_interest_area", new Achievement_Meta()));
            Assert.AreEqual("Adventurous Journey", ApprovalNames.AchievementName("adventurous_journey", null));
        }
    }
}
