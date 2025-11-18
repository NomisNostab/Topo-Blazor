using Newtonsoft.Json;
using Syncfusion.XlsIO;
using System.Globalization;
using System.Xml.Linq;
using Topo.Model.Approvals;
using Topo.Model.Members;
using Topo.Model.ReportGeneration;
using TopoReportFunction;

namespace TopoReportFunctionTest
{
    [TestClass]
    public class Approvals
    {
        private ReportGenerationRequest _reportGenerationRequest;
        private Function _function;

        [TestInitialize]
        public void SetUp()
        {
            SetEnvironment.SetSyncfusionLicense();
            _function = new Function();
            _reportGenerationRequest = new ReportGenerationRequest()
            {
                ReportType = ReportType.Approvals,
                GroupName = "Epping Scout Group",
                Section = "scout",
                UnitName = "Scout Unit",
                OutputType = OutputType.PDF,
                ReportData = ""
            };
            if (!Directory.Exists("TestResults"))
                Directory.CreateDirectory("TestResults");
        }

        [TestMethod]
        public void Approvals_ToPDF()
        {
            var proxyRequest = new Amazon.Lambda.APIGatewayEvents.APIGatewayProxyRequest();
            proxyRequest.Body = JsonConvert.SerializeObject(_reportGenerationRequest);
            var functionResult = _function.FunctionHandler(proxyRequest, null);
            //Convert Base64String into PDF document
            byte[] bytes = Convert.FromBase64String(functionResult);
            FileStream fileStream = new FileStream(@"TestResults\AdditionalAwards_ToPDF.pdf", FileMode.Create);
            BinaryWriter writer = new BinaryWriter(fileStream);
            writer.Write(bytes, 0, bytes.Length);
            writer.Close();
        }

        [TestMethod]
        public void ApprovalsParsing()
        {
            string pendingApprovalsString = "{\"results\": []}";
            string finalisedApprovalsString = "{\"results\": []}";
            string memberListString = "[]";
            var pendingApprovals = JsonConvert.DeserializeObject<GetApprovalsResultModel>(pendingApprovalsString);
            var finalisedApprovals = JsonConvert.DeserializeObject<GetApprovalsResultModel>(finalisedApprovalsString);
            var approvals = finalisedApprovals.results.Concat(pendingApprovals.results).ToList();
            var achievementApprovals = approvals.Where(a => a.submission.date > new DateTime(2024, 1, 1)).GroupBy(r => r.achievement.id).ToList();
            var memberList = JsonConvert.DeserializeObject<List<MemberListModel>>(memberListString);
            List<ApprovalsListModel> approvalList = new List<ApprovalsListModel>();
            foreach (var result in achievementApprovals)
            {
                string approvalDate = "";
                string awardedDate = "";
                ApprovalsListModel approvalItem = new ApprovalsListModel();
                foreach (var item in result)
                {
                    Console.Write($"Type:{item.achievement.type} ");
                    approvalItem.achievement_type = item.achievement.type;
                    approvalItem.achievement_id = item.achievement.id;
                    switch (item.achievement.type)
                    {
                        case "special_interest_area":
                            Console.Write($"Area:{item.achievement.achievement_meta?.sia_area} ");
                            approvalItem.achievement_name = item.achievement.achievement_meta?.sia_area.Replace("_", " ");
                            break;
                        case "milestone":
                            Console.Write($"Stage:{item.achievement.achievement_meta?.stage} ");
                            approvalItem.achievement_name = $"milestone {item.achievement.achievement_meta?.stage}";
                            break;
                        case "outdoor_adventure_skill":
                            Console.Write($"Stream:{item.achievement.achievement_meta?.stream} Stage:{item.achievement.achievement_meta?.stage} ");
                            approvalItem.achievement_name = item.achievement.achievement_meta?.stream == item.achievement.achievement_meta?.branch
                                    ? $"{item.achievement.achievement_meta?.stream} {item.achievement.achievement_meta?.stage}"
                                    : $"{item.achievement.achievement_meta?.stream} {item.achievement.achievement_meta?.branch} {item.achievement.achievement_meta?.stage}";
                            break;
                        case "intro_section":
                            approvalItem.achievement_name = "intro to section";
                            break;
                        case "intro_scouting":
                            approvalItem.achievement_name = "intro to scouting";
                            break;
                        default:
                            approvalItem.achievement_name = item.achievement.type.Replace("_", " ");
                            break;
                    }
                    TextInfo myTI = new CultureInfo("en-AU", false).TextInfo;
                    approvalItem.achievement_name = myTI.ToTitleCase(approvalItem.achievement_name);

                    approvalItem.member_id = item.member.id;
                    approvalItem.member_first_name = item.member.first_name;
                    approvalItem.member_last_name = item.member.last_name;
                    Console.Write($"Name:{item.member.first_name} {item.member.last_name} ");

                    approvalItem.submission_type = myTI.ToTitleCase(item.submission.type);
                    if (item.submission.type == "review")
                    {
                        approvalItem.submission_date = item.submission.date;
                    }
                    approvalItem.submission_status = myTI.ToTitleCase(item.submission.status);
                    approvalItem.submission_outcome = myTI.ToTitleCase(item.submission.outcome);

                    Console.Write($"Submission Type:{item.submission.type} Status:{item.submission.status} Outcome:{item.submission.outcome} ");
                    int actionCount = 0;
                    foreach (var action in item.submission.actioned_by)
                    {
                        if (action.outcome == "approved")
                        {
                            actionCount++;
                        }
                        Console.Write($"Action outcome:{action.outcome} Action time:{action.time} Action date:{action.date_awarded} ");
                        switch (item.submission.type)
                        {
                            case "review":
                                if (actionCount == 2)
                                {
                                    approvalDate = action.time.ToShortDateString();
                                    approvalItem.awarded_date = action.time;
                                }
                                break;
                            case "award":
                                // need to check if action.member_id is a leader
                                bool isLeader = memberList.Where(m => m.id == action.member_id).FirstOrDefault().isAdultLeader == 1;
                                if (isLeader)
                                {
                                    awardedDate = string.IsNullOrEmpty(action.date_awarded) ? action.time.ToShortDateString() : action.date_awarded;
                                    approvalItem.presented_date = action.time;
                                    if (!string.IsNullOrEmpty(action.date_awarded))
                                    {
                                        approvalItem.presented_date = DateTime.ParseExact(action.date_awarded, "yyyy-MM-dd", CultureInfo.InvariantCulture);
                                    }
                                }
                                break;
                        }
                    }
                    Console.WriteLine();
                }
                Console.WriteLine($"Approved: {approvalDate} Awarded:{awardedDate}");
                approvalList.Add(approvalItem);
            }
        }
    }
}
