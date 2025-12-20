using Newtonsoft.Json;
using Topo.Model.ReportGeneration;
using TopoReportFunction;

namespace TopoReportFunctionTest
{
    [TestClass]
    public class Progress
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
                ReportType = ReportType.PersonalProgress,
                GroupName = "Epping Scout Group",
                Section = "scout",
                UnitName = "Scout Unit",
                OutputType = OutputType.PDF,
                ReportData = "{\"Member\":{\"id\":\"a83d4ffd-ba2d-3952-8ea2-294f362e763c\",\"member_number\":\"99999999\",\"first_name\":\"Joe\",\"last_name\":\"Blow\",\"status\":\"active\",\"age\":\"13y 3m\",\"unit_council\":false,\"patrol_name\":\"Drop Bear (Green)\",\"patrol_duty\":\"\",\"patrol_order\":3,\"isAdultLeader\":0,\"selected\":false,\"isEligibleJamboree\":true},\"GroupName\":\"(Scouts NSW - Epping Scout Group)\",\"IntroToScoutingDate\":\"28/09/25\",\"IntroToSectionDate\":\"28/09/2025\",\"Milestones\":[{\"Milestone\":1,\"Status\":\"In Progress\",\"Awarded\":\"0001-01-01T00:00:00\",\"ParticipateLogs\":[{\"ChallengeArea\":\"creative\",\"EventName\":\"Navigation\",\"EventDate\":\"2025-07-30T19:00:00+10:00\",\"ChallengeAreaAbbrev\":\"(Cr)\"},{\"ChallengeArea\":\"creative\",\"EventName\":\"Science Night\",\"EventDate\":\"2025-08-20T19:00:00+10:00\",\"ChallengeAreaAbbrev\":\"(Cr)\"},{\"ChallengeArea\":\"creative\",\"EventName\":\"Raft making/floating\",\"EventDate\":\"2025-11-05T18:30:00+11:00\",\"ChallengeAreaAbbrev\":\"(Cr)\"},{\"ChallengeArea\":\"creative\",\"EventName\":\"Cork Gun Night\",\"EventDate\":\"2025-11-19T19:00:00+11:00\",\"ChallengeAreaAbbrev\":\"(Cr)\"},{\"ChallengeArea\":\"community\",\"EventName\":\"Hall Centennial Sausage Sizzle\",\"EventDate\":\"2025-07-26T10:00:00+10:00\",\"ChallengeAreaAbbrev\":\"(Co)\"},{\"ChallengeArea\":\"community\",\"EventName\":\"Group Camp Packing\",\"EventDate\":\"2025-11-26T19:00:00+11:00\",\"ChallengeAreaAbbrev\":\"(Co)\"},{\"ChallengeArea\":\"community\",\"EventName\":\"Beach Swim - Rescues/Fun\",\"EventDate\":\"2025-12-03T19:00:00+11:00\",\"ChallengeAreaAbbrev\":\"(Co)\"},{\"ChallengeArea\":\"personal_growth\",\"EventName\":\"Bushwalking Thoery\",\"EventDate\":\"2025-07-23T19:00:00+10:00\",\"ChallengeAreaAbbrev\":\"(PG)\"},{\"ChallengeArea\":\"personal_growth\",\"EventName\":\"Special Interest Area - Setting\",\"EventDate\":\"2025-08-06T19:00:00+10:00\",\"ChallengeAreaAbbrev\":\"(PG)\"},{\"ChallengeArea\":\"personal_growth\",\"EventName\":\"Bike Maintenance and Theory \",\"EventDate\":\"2025-08-27T19:00:00+10:00\",\"ChallengeAreaAbbrev\":\"(PG)\"},{\"ChallengeArea\":\"personal_growth\",\"EventName\":\"Abseiling Vertical Theory\",\"EventDate\":\"2025-09-10T19:00:00+10:00\",\"ChallengeAreaAbbrev\":\"(PG)\"},{\"ChallengeArea\":\"personal_growth\",\"EventName\":\"Water Safety\",\"EventDate\":\"2025-10-15T19:00:00+11:00\",\"ChallengeAreaAbbrev\":\"(PG)\"},{\"ChallengeArea\":\"personal_growth\",\"EventName\":\"Canoeing - Rescues\",\"EventDate\":\"2025-12-10T18:30:00+11:00\",\"ChallengeAreaAbbrev\":\"(PG)\"},{\"ChallengeArea\":\"outdoors\",\"EventName\":\"Bushwalk\",\"EventDate\":\"2025-08-13T19:00:00+10:00\",\"ChallengeAreaAbbrev\":\"(O)\"},{\"ChallengeArea\":\"outdoors\",\"EventName\":\"Abseiling Day\",\"EventDate\":\"2025-09-14T12:00:00+10:00\",\"ChallengeAreaAbbrev\":\"(O)\"},{\"ChallengeArea\":\"outdoors\",\"EventName\":\"Abseiling\",\"EventDate\":\"2025-09-17T19:00:00+10:00\",\"ChallengeAreaAbbrev\":\"(O)\"},{\"ChallengeArea\":\"outdoors\",\"EventName\":\"Rock Climbing\",\"EventDate\":\"2025-09-24T19:00:00+10:00\",\"ChallengeAreaAbbrev\":\"(O)\"},{\"ChallengeArea\":\"outdoors\",\"EventName\":\"Waterway Safety\",\"EventDate\":\"2025-10-29T19:00:00+11:00\",\"ChallengeAreaAbbrev\":\"(O)\"},{\"ChallengeArea\":\"outdoors\",\"EventName\":\"Canoeing - Journey\",\"EventDate\":\"2025-11-12T18:30:00+11:00\",\"ChallengeAreaAbbrev\":\"(O)\"}],\"AssistLogs\":[],\"LeadLogs\":[]}],\"OASSummaries\":[{\"Stream\":\"aquatics\",\"Stage\":1,\"Awarded\":\"0001-01-01T00:00:00\",\"Approved\":\"2025-11-12T22:02:50.241169+11:00\",\"Section\":\"s\",\"Template\":\"oas/aquatics/1\"},{\"Stream\":\"aquatics\",\"Stage\":2,\"Awarded\":\"0001-01-01T00:00:00\",\"Approved\":\"0001-01-01T00:00:00\",\"Section\":\"s\",\"Template\":\"oas/aquatics/2\"},{\"Stream\":\"aquatics\",\"Stage\":3,\"Awarded\":\"0001-01-01T00:00:00\",\"Approved\":\"0001-01-01T00:00:00\",\"Section\":\"s\",\"Template\":\"oas/aquatics/3\"},{\"Stream\":\"aquatics\",\"Stage\":4,\"Awarded\":\"0001-01-01T00:00:00\",\"Approved\":\"0001-01-01T00:00:00\",\"Section\":\"s\",\"Template\":\"oas/aquatics/life-saving/4\"},{\"Stream\":\"boating\",\"Stage\":1,\"Awarded\":\"0001-01-01T00:00:00\",\"Approved\":\"0001-01-01T00:00:00\",\"Section\":\"s\",\"Template\":\"oas/boating/1\"},{\"Stream\":\"boating\",\"Stage\":2,\"Awarded\":\"0001-01-01T00:00:00\",\"Approved\":\"0001-01-01T00:00:00\",\"Section\":\"s\",\"Template\":\"oas/boating/2\"},{\"Stream\":\"boating\",\"Stage\":3,\"Awarded\":\"0001-01-01T00:00:00\",\"Approved\":\"0001-01-01T00:00:00\",\"Section\":\"s\",\"Template\":\"oas/boating/3\"},{\"Stream\":\"boating\",\"Stage\":4,\"Awarded\":\"0001-01-01T00:00:00\",\"Approved\":\"0001-01-01T00:00:00\",\"Section\":\"s\",\"Template\":\"oas/boating/sailing/4\"},{\"Stream\":\"bushwalking\",\"Stage\":1,\"Awarded\":\"2025-11-05T22:01:36.529043+11:00\",\"Approved\":\"2025-11-12T22:01:36.529043+11:00\",\"Section\":\"s\",\"Template\":\"oas/bushwalking/1\"},{\"Stream\":\"bushwalking\",\"Stage\":2,\"Awarded\":\"0001-01-01T00:00:00\",\"Approved\":\"2025-11-12T22:02:06.666019+11:00\",\"Section\":\"s\",\"Template\":\"oas/bushwalking/2\"},{\"Stream\":\"bushwalking\",\"Stage\":3,\"Awarded\":\"0001-01-01T00:00:00\",\"Approved\":\"2025-10-22T21:40:51.862481+11:00\",\"Section\":\"s\",\"Template\":\"oas/bushwalking/3\"},{\"Stream\":\"camping\",\"Stage\":1,\"Awarded\":\"0001-01-01T00:00:00\",\"Approved\":\"2025-12-03T20:47:29.745389+11:00\",\"Section\":\"s\",\"Template\":\"oas/camping/1\"},{\"Stream\":\"cycling\",\"Stage\":1,\"Awarded\":\"0001-01-01T00:00:00\",\"Approved\":\"0001-01-01T00:00:00\",\"Section\":\"s\",\"Template\":\"oas/cycling/1\"},{\"Stream\":\"cycling\",\"Stage\":2,\"Awarded\":\"0001-01-01T00:00:00\",\"Approved\":\"0001-01-01T00:00:00\",\"Section\":\"s\",\"Template\":\"oas/cycling/2\"},{\"Stream\":\"cycling\",\"Stage\":3,\"Awarded\":\"0001-01-01T00:00:00\",\"Approved\":\"0001-01-01T00:00:00\",\"Section\":\"s\",\"Template\":\"oas/cycling/3\"},{\"Stream\":\"vertical\",\"Stage\":1,\"Awarded\":\"0001-01-01T00:00:00\",\"Approved\":\"0001-01-01T00:00:00\",\"Section\":\"s\",\"Template\":\"oas/vertical/1\"},{\"Stream\":\"vertical\",\"Stage\":2,\"Awarded\":\"0001-01-01T00:00:00\",\"Approved\":\"0001-01-01T00:00:00\",\"Section\":\"s\",\"Template\":\"oas/vertical/2\"}],\"Stats\":{\"OasProgressions\":0,\"KmsHiked\":6,\"NightsCamped\":3,\"NightsCampedInSection\":3},\"SIASummaries\":[],\"PeakAward\":{\"PersonalDevelopmentCourse\":\"\",\"AdventurousJourney\":\"\",\"PersonalReflection\":\"\",\"Awarded\":\"\"},\"DisableOAS\":null,\"DisableCoreOAS\":null}"
            };
            if (!Directory.Exists("TestResults"))
                Directory.CreateDirectory("TestResults");
        }

        [TestMethod]
        public void Progress_ToPDF()
        {
            var proxyRequest = new Amazon.Lambda.APIGatewayEvents.APIGatewayProxyRequest();
            proxyRequest.Body = JsonConvert.SerializeObject(_reportGenerationRequest);
            var functionResult = _function.FunctionHandler(proxyRequest, null);
            //Convert Base64String into PDF document
            byte[] bytes = Convert.FromBase64String(functionResult);
            FileStream fileStream = new FileStream(@"TestResults\Progress_ToPDF.pdf", FileMode.Create);
            BinaryWriter writer = new BinaryWriter(fileStream);
            writer.Write(bytes, 0, bytes.Length);
            writer.Close();
        }

        [TestMethod]
        public void Progress_ToExcel()
        {
            var proxyRequest = new Amazon.Lambda.APIGatewayEvents.APIGatewayProxyRequest();
            _reportGenerationRequest.OutputType = OutputType.Excel;
            proxyRequest.Body = JsonConvert.SerializeObject(_reportGenerationRequest);
            var functionResult = _function.FunctionHandler(proxyRequest, null);
            //Convert Base64String into PDF document
            byte[] bytes = Convert.FromBase64String(functionResult);
            FileStream fileStream = new FileStream(@"TestResults\Progress_ToExcel.xlsx", FileMode.Create);
            BinaryWriter writer = new BinaryWriter(fileStream);
            writer.Write(bytes, 0, bytes.Length);
            writer.Close();
        }

    }
}
