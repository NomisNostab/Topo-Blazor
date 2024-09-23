using System.Globalization;
using Topo.Model.Members;

namespace Topo.Services
{
    public interface IMembersService
    {
        public Task<List<MemberListModel>> GetMembersAsync(string unitId);
        public void ClearMemberCache(string unitId);
        public Task<MemberListModel> GetMember(string unitId, string memberId);
        public Task<string> GetMemberLastName(string unitId, string memberId);
    }

    public class MembersService : IMembersService
    {
        private readonly StorageService _storageService;
        private readonly ITerrainAPIService _terrainAPIService;

        public MembersService(StorageService storageService, ITerrainAPIService terrainAPIService)
        {
            _storageService = storageService;
            _terrainAPIService = terrainAPIService;
        }

        public void ClearMemberCache(string unitId)
        {
            var cachedMembersList = _storageService.CachedMembers.Where(cm => cm.Key == unitId).FirstOrDefault();
            _storageService.CachedMembers.Remove(cachedMembersList);
        }

        public async Task<List<MemberListModel>> GetMembersAsync(string unitId)
        {
            var cachedMembersList = _storageService.CachedMembers.Where(cm => cm.Key == unitId).FirstOrDefault().Value;
            if (cachedMembersList != null)
                return cachedMembersList;

            var getMembersResultModel = await _terrainAPIService.GetMembersAsync(unitId ?? "");
            var memberList = new List<MemberListModel>();
            if (_storageService.IsYouthMember)
            {
                if (_storageService.GetProfilesResult != null && _storageService.GetProfilesResult.profiles != null && _storageService.GetProfilesResult.profiles.Length > 0)
                {
                    var usernameSplit = _storageService.GetProfilesResult.username.Split("-");
                    var profile = _storageService.GetProfilesResult.profiles[0];
                    var nameSplit = profile.member.name.Split(" ");
                    memberList.Add(
                            new MemberListModel
                            {
                                id = profile.member.id,
                                member_number = usernameSplit.Length > 1 ? usernameSplit[1] : "",
                                first_name = nameSplit.Length > 0 ? nameSplit[0] : "",
                                last_name = nameSplit.Length > 1 ? nameSplit[1] + (nameSplit.Length > 2 ? $"{nameSplit[2]} " : "") : "",
                                isAdultLeader = 0
                            });

                }
            }
            else
            {
                if (getMembersResultModel != null && getMembersResultModel.results != null)
                {
                    memberList = getMembersResultModel.results
                        .Select(m => new MemberListModel
                        {
                            id = m.id,
                            member_number = m.member_number,
                            first_name = m.first_name,
                            last_name = m.last_name,
                            age = GetAgeFromBirthdate(m.date_of_birth),
                            unit_council = m.unit.unit_council,
                            patrol_name = m.patrol == null ? "" : m.patrol.name,
                            patrol_duty = GetPatrolDuty(m.unit.duty, m.patrol?.duty ?? ""),
                            patrol_order = GetPatrolOrder(m.unit.duty, m.patrol?.duty ?? ""),
                            isAdultLeader = m.unit.duty == "adult_leader" ? 1 : 0,
                            status = m.status,
                            isEligibleJamboree = GetJamboreeEligibilityFromBirthdate(m.date_of_birth)
                        })
                        .ToList();
                    _storageService.CachedMembers.Add(new KeyValuePair<string, List<MemberListModel>>(unitId, memberList));
                }
            }
            return memberList;
        }

        public async Task<MemberListModel> GetMember(string unitId, string memberId)
        {
            var members = await GetMembersAsync(unitId);
            var member = members.Where(m => m.id == memberId).FirstOrDefault();
            return member ?? new MemberListModel();
        }

        public async Task<string> GetMemberLastName(string unitId, string memberId)
        {
            var members = await GetMembersAsync(unitId);
            var member = members.Where(m => m.id == memberId).FirstOrDefault();
            if (member == null || string.IsNullOrEmpty(member.last_name))
            {
                return "";
            }
            if (_storageService.SuppressLastName)
            {
                var lastName = member.last_name.Substring(0, 1).ToUpper();
                return lastName;
            }
            else
            {
                return member.last_name;
            }
        }

        private string GetAgeFromBirthdate(string dateOfBirth)
        {
            var birthday = DateTime.ParseExact(dateOfBirth, "yyyy-MM-dd", CultureInfo.InvariantCulture); // Date in AU format
            DateTime now = DateTime.Today;
            int months = now.Month - birthday.Month;
            int years = now.Year - birthday.Year;

            if (now.Day < birthday.Day)
            {
                months--;
            }

            if (months < 0)
            {
                years--;
                months += 12;
            }
            return $"{years}y {months}m";
        }

        private bool GetJamboreeEligibilityFromBirthdate(string dateOfBirth)
        {
            var birthday = DateTime.ParseExact(dateOfBirth, "yyyy-MM-dd", CultureInfo.InvariantCulture); // Date in AU format
            var scoutMaxCutOffDate = new DateTime(2010, 1, 6);
            var scoutMinCutOffDate = new DateTime(2014, 1, 6);
            var venturerMaxCutOffDate = new DateTime(2008, 1, 6);
            if (birthday > scoutMaxCutOffDate && birthday < scoutMinCutOffDate)
            {
                return true ;
            }
            if (birthday > venturerMaxCutOffDate && birthday < scoutMaxCutOffDate)
            {
                return true;
            }
            return false;
        }

        private string GetPatrolDuty(string unitDuty, string patrolDuty)
        {
            if (unitDuty == "adult_leader")
                return "SL";
            if (unitDuty == "unit_leader")
                return "UL";
            if (patrolDuty == "assistant_patrol_leader")
                return "APL";
            if (patrolDuty == "patrol_leader")
                return "PL";
            return "";
        }

        private int GetPatrolOrder(string unitDuty, string patrolDuty)
        {
            if (unitDuty == "unit_leader")
                return 0;
            if (patrolDuty == "assistant_patrol_leader")
                return 2;
            if (patrolDuty == "patrol_leader")
                return 1;
            return 3;
        }


    }
}
