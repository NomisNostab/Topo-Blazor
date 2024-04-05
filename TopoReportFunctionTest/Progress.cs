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
            _function = new Function();
            _reportGenerationRequest = new ReportGenerationRequest()
            {
                ReportType = ReportType.PersonalProgress,
                GroupName = "Epping Scout Group",
                Section = "scout",
                UnitName = "Scout Unit",
                OutputType = OutputType.PDF,
                ReportData = "{\"Member\":{\"id\":\"abc9a1eb-5f49-333a-8899-d7c338edf80a\",\"member_number\":\"99999999\",\"first_name\":\"First Name\",\"last_name\":\"Last Name\",\"status\":\"active\",\"age\":\"8y 5m\",\"unit_council\":false,\"patrol_name\":\"Joeys - Red\",\"patrol_duty\":\"\",\"patrol_order\":3,\"isAdultLeader\":0,\"selected\":true,\"isEligibleJamboree\":false},\"GroupName\":\"(1ST GLEN IRIS)\",\"IntroToScoutingDate\":\"08/05/23\",\"IntroToSectionDate\":\"08/05/2023\",\"Milestones\":[{\"Milestone\":1,\"Status\":\"Not Required\",\"Awarded\":\"0001-01-01T00:00:00\",\"ParticipateLogs\":[],\"AssistLogs\":[],\"LeadLogs\":[]},{\"Milestone\":2,\"Status\":\"Not Required\",\"Awarded\":\"0001-01-01T00:00:00\",\"ParticipateLogs\":[],\"AssistLogs\":[],\"LeadLogs\":[]},{\"Milestone\":3,\"Status\":\"Awarded\",\"Awarded\":\"2024-02-12T20:48:42.098249+11:00\",\"ParticipateLogs\":[{\"ChallengeArea\":\"community\",\"EventName\":\"Imported\",\"EventDate\":\"2024-02-12T20:48:42.098249+11:00\"},{\"ChallengeArea\":\"outdoors\",\"EventName\":\"Imported\",\"EventDate\":\"2024-02-12T20:48:42.098249+11:00\"},{\"ChallengeArea\":\"personal_growth\",\"EventName\":\"Prepare for camping & first aid\",\"EventDate\":\"2023-05-29T17:45:00+10:00\"},{\"ChallengeArea\":\"personal_growth\",\"EventName\":\"Movie & Popcorn \",\"EventDate\":\"2023-07-10T17:45:00+10:00\"},{\"ChallengeArea\":\"personal_growth\",\"EventName\":\"Construction night\",\"EventDate\":\"2023-08-14T17:45:00+10:00\"},{\"ChallengeArea\":\"personal_growth\",\"EventName\":\"First night back\",\"EventDate\":\"2023-10-09T17:45:00+11:00\"},{\"ChallengeArea\":\"outdoors\",\"EventName\":\"Investiture & Campfire\",\"EventDate\":\"2023-06-19T17:45:00+10:00\"},{\"ChallengeArea\":\"outdoors\",\"EventName\":\"Night hike\",\"EventDate\":\"2023-08-28T17:45:00+10:00\"},{\"ChallengeArea\":\"outdoors\",\"EventName\":\"Valley Reserve with B Pack\",\"EventDate\":\"2023-10-17T18:30:00+11:00\"},{\"ChallengeArea\":\"community\",\"EventName\":\"Joint night with B Pack\",\"EventDate\":\"2023-06-13T18:30:00+10:00\"},{\"ChallengeArea\":\"community\",\"EventName\":\"Public Transport night\",\"EventDate\":\"2023-07-17T17:45:00+10:00\"},{\"ChallengeArea\":\"community\",\"EventName\":\"Flight night\",\"EventDate\":\"2023-09-04T17:45:00+10:00\"},{\"ChallengeArea\":\"creative\",\"EventName\":\"Glow night\",\"EventDate\":\"2023-07-24T17:45:00+10:00\"},{\"ChallengeArea\":\"creative\",\"EventName\":\"Cooking night\",\"EventDate\":\"2023-08-07T17:45:00+10:00\"},{\"ChallengeArea\":\"creative\",\"EventName\":\"Craft night\",\"EventDate\":\"2023-10-30T17:45:00+11:00\"},{\"ChallengeArea\":\"creative\",\"EventName\":\"Cooking night\",\"EventDate\":\"2023-11-13T17:45:00+11:00\"}],\"AssistLogs\":[{\"ChallengeArea\":\"creative\",\"EventName\":\"Imported\",\"EventDate\":\"2024-02-12T20:48:42.098249+11:00\"},{\"ChallengeArea\":\"community\",\"EventName\":\"Community helpers night\",\"EventDate\":\"2023-05-22T17:45:00+10:00\"},{\"ChallengeArea\":\"personal_growth\",\"EventName\":\"Environmental Warriors Weekend camp\",\"EventDate\":\"2023-06-02T17:00:00+10:00\"},{\"ChallengeArea\":\"personal_growth\",\"EventName\":\"Science night\",\"EventDate\":\"2023-12-04T17:45:00+11:00\"}],\"LeadLogs\":[{\"ChallengeArea\":\"personal_growth\",\"EventName\":\"Bushcraft night\",\"EventDate\":\"2023-07-31T17:45:00+10:00\"},{\"ChallengeArea\":\"creative\",\"EventName\":\"Paper fashion\",\"EventDate\":\"2023-11-27T17:45:00+11:00\"},{\"ChallengeArea\":\"creative\",\"EventName\":\"Wet & Messy night\",\"EventDate\":\"2023-12-11T17:45:00+11:00\"},{\"ChallengeArea\":\"personal_growth\",\"EventName\":\"Logbook for sleepover\",\"EventDate\":\"2023-12-16T17:00:00+11:00\"}]}],\"OASSummaries\":[{\"Stream\":\"alpine\",\"Stage\":1,\"Awarded\":\"2023-07-11T23:50:47.256784+10:00\",\"Section\":\"j\",\"Template\":\"oas/alpine/1\"},{\"Stream\":\"aquatics\",\"Stage\":1,\"Awarded\":\"2024-03-19T23:38:32.60916+11:00\",\"Section\":\"j\",\"Template\":\"oas/aquatics/1\"},{\"Stream\":\"bushcraft\",\"Stage\":1,\"Awarded\":\"2024-02-12T20:43:27.426831+11:00\",\"Section\":\"j\",\"Template\":\"oas/bushcraft/1\"},{\"Stream\":\"bushcraft\",\"Stage\":2,\"Awarded\":\"0001-01-01T00:00:00\",\"Section\":\"j\",\"Template\":\"oas/bushcraft/2\"},{\"Stream\":\"bushwalking\",\"Stage\":1,\"Awarded\":\"2023-11-25T12:01:12.813272+11:00\",\"Section\":\"j\",\"Template\":\"oas/bushwalking/1\"},{\"Stream\":\"bushwalking\",\"Stage\":2,\"Awarded\":\"0001-01-01T00:00:00\",\"Section\":\"j\",\"Template\":\"oas/bushwalking/2\"},{\"Stream\":\"camping\",\"Stage\":1,\"Awarded\":\"2023-08-14T09:37:23.979015+10:00\",\"Section\":\"j\",\"Template\":\"oas/camping/1\"},{\"Stream\":\"camping\",\"Stage\":2,\"Awarded\":\"0001-01-01T00:00:00\",\"Section\":\"j\",\"Template\":\"oas/camping/2\"},{\"Stream\":\"cycling\",\"Stage\":1,\"Awarded\":\"2024-03-26T15:23:41.803336+11:00\",\"Section\":\"j\",\"Template\":\"oas/cycling/1\"},{\"Stream\":\"vertical\",\"Stage\":1,\"Awarded\":\"0001-01-01T00:00:00\",\"Section\":\"j\",\"Template\":\"oas/vertical/1\"}],\"Stats\":{\"OasProgressions\":6,\"KmsHiked\":25,\"NightsCamped\":3,\"NightsCampedInSection\":3},\"SIASummaries\":[{\"Area\":\"Art Literature\",\"Project\":\"Play the cello\",\"Status\":\"Awarded\",\"Awarded\":\"2023-11-12T19:19:47.877601+11:00\"},{\"Area\":\"Adventure Sport\",\"Project\":\"180o ski jump\",\"Status\":\"Awarded\",\"Awarded\":\"2023-09-07T16:46:12.523343+10:00\"},{\"Area\":\"Growth Development\",\"Project\":\"Learn how to make pasta from scratch\",\"Status\":\"Awarded\",\"Awarded\":\"2023-11-12T19:19:59.177279+11:00\"},{\"Area\":\"Environment\",\"Project\":\"Operation Clean Up\",\"Status\":\"Awarded\",\"Awarded\":\"2024-03-19T23:41:20.492186+11:00\"},{\"Area\":\"Stem Innovation\",\"Project\":\"Build a space explorers robotics kit\",\"Status\":\"Awarded\",\"Awarded\":\"2024-02-12T20:44:24.975107+11:00\"},{\"Area\":\"Better World\",\"Project\":\"Plant & grow tomatoes\",\"Status\":\"Awarded\",\"Awarded\":\"2024-02-12T20:48:31.764077+11:00\"},{\"Area\":\"Art Literature\",\"Project\":\"Sew a baby bib\",\"Status\":\"Awarded\",\"Awarded\":\"2024-02-12T20:50:07.197189+11:00\"}],\"PeakAward\":{\"PersonalDevelopmentCourse\":\"\",\"AdventurousJourney\":\"25/11/23\",\"PersonalReflection\":\"12/02/24\",\"Awarded\":\"24/02/24\"}}"
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
