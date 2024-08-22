using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using System.Globalization;
using Topo.Model.Program;
using System;

namespace Topo.Services
{
    public interface IProgramService
    {
        public Task<Dictionary<string, string>> GetCalendars();
        public Task SetCalendar(string calendarId);
        public Task ResetCalendar();
        public Task<List<EventListModel>> GetEventsForDates(DateTime fromDate, DateTime toDate);
        public Task<EventListModel> GetAttendanceForEvent(string eventId);
        public Task<AttendanceReportModel> GenerateAttendanceReportData(DateTime fromDate, DateTime toDate, string selectedCalendar, string groupCalendar);
    }

    public class ProgramService : IProgramService
    {
        private readonly StorageService _storageService;
        private readonly ITerrainAPIService _terrainAPIService;
        private readonly IMembersService _memberService;

        private GetCalendarsResultModel getCalendarsResultModel;

        public ProgramService(StorageService storageService, ITerrainAPIService terrainAPIService, IMembersService membersService)
        {
            _storageService = storageService;
            _terrainAPIService = terrainAPIService;
            _memberService = membersService;
        }

        private string GetUser()
        {
            var userId = "";
            if (_storageService.GetProfilesResult != null && _storageService.GetProfilesResult.profiles != null && _storageService.GetProfilesResult.profiles.Length > 0)
            {
                userId = _storageService.GetProfilesResult.profiles[0].member?.id;
            }
            return userId ?? "";
        }

        public async Task<Dictionary<string, string>> GetCalendars()
        {
            getCalendarsResultModel = await _terrainAPIService.GetCalendarsAsync(GetUser());
            if (getCalendarsResultModel != null && getCalendarsResultModel.own_calendars != null)
            {
                var calendars = getCalendarsResultModel.own_calendars.Where(c => c.type == "unit")
                    .ToDictionary(x => x.id, x => x.title);
                return calendars;
            }
            return new Dictionary<string, string>();
        }

        public async Task SetCalendar(string calendarId)
        {
            //Deep copy calendar to allow reset
            var serialisedCalendar = JsonConvert.SerializeObject(getCalendarsResultModel);
            var calendars = JsonConvert.DeserializeObject<GetCalendarsResultModel>(serialisedCalendar);

            foreach (var calendar in calendars.own_calendars)
            {
                calendar.selected = calendar.id == calendarId;
            }

            foreach (var calendar in calendars.other_calendars)
            {
                calendar.selected = false;
            }

            await _terrainAPIService.PutCalendarsAsync(GetUser(), calendars);
        }

        public async Task ResetCalendar()
        {
            await _terrainAPIService.PutCalendarsAsync(GetUser(), getCalendarsResultModel);
        }

        public async Task<List<EventListModel>> GetEventsForDates(DateTime fromDate, DateTime toDate)
        {
            TextInfo myTI = new CultureInfo("en-AU", false).TextInfo;
            var getEventsResultModel = await _terrainAPIService.GetEventsAsync(GetUser(), fromDate.AddDays(-1), toDate);
            if (getEventsResultModel != null && getEventsResultModel.results != null)
            {
                var eventList = new List<EventListModel>();
                foreach (var eventResult in getEventsResultModel.results)
                {
                    var getEventResultModel = await _terrainAPIService.GetEventAsync(eventResult.id);
                    var leads = getEventResultModel.attendance.leader_members.Select(a => string.Concat(a.first_name, " ", a.last_name.AsSpan(0, 1)));
                    var leadNames = string.Join(", ", leads);
                    var assists = getEventResultModel.attendance.assistant_members.Select(a => string.Concat(a.first_name, " ", a.last_name.AsSpan(0, 1)));
                    var assistNames = string.Join(", ", assists);
                    var organisers = getEventResultModel.organisers.Select(a => string.Concat(a.first_name, " ", a.last_name.AsSpan(0, 1)));
                    var organiserNames = string.Join(", ", organisers);
                    eventList.Add(new EventListModel()
                    {
                        Id = eventResult.id,
                        EventName = eventResult.title,
                        StartDateTime = eventResult.start_datetime,
                        StartDateTimeDisplay = eventResult.start_datetime.ToString("dd/MM/yy HH:mm"),
                        EndDateTime = eventResult.end_datetime,
                        EndDateTimeDisplay = eventResult.end_datetime.ToString("dd/MM/yy HH:mm"),
                        StartFinishDisplay = eventResult.start_datetime.Date == eventResult.end_datetime.Date
                            ? $"{eventResult.start_datetime:h:mm tt} - {eventResult.end_datetime:h:mm tt}"
                            : $"{eventResult.start_datetime:ddd h:mm tt} - {eventResult.end_datetime:ddd h:mm tt}",
                        ChallengeArea = myTI.ToTitleCase(eventResult.challenge_area.Replace("_", " ").Replace("personal ", "")),
                        EventStatus = myTI.ToTitleCase(eventResult.status),
                        IsUnitEvent = eventResult.invitee_type == "unit",
                        EventDisplay = $"{eventResult.title} {eventResult.start_datetime.ToShortDateString()}",
                        Organiser = organiserNames,
                        Lead = leadNames,
                        Assist = assistNames,
                        Location = getEventResultModel?.location ?? ""
                    });

                }
                return eventList;
            }
            return new List<EventListModel>();
        }

        public async Task<EventListModel> GetAttendanceForEvent(string eventId)
        {
            var eventListModel = new EventListModel();
            var getEventResultModel = await _terrainAPIService.GetEventAsync(eventId);
            var eventAttendance = new List<EventAttendance>();
            var members = await _memberService.GetMembersAsync(_storageService.UnitId);
            foreach (var member in members)
            {
                eventAttendance.Add(new EventAttendance
                {
                    id = member.id,
                    first_name = member.first_name,
                    last_name = member.last_name,
                    member_number = member.member_number,
                    patrol_name = member.patrol_name,
                    isAdultMember = member.isAdultLeader == 1,
                    attended = false,
                    pal = ""
                });
            }

            eventListModel.Id = eventId;
            eventListModel.EventName = getEventResultModel?.title ?? "Event not found";
            eventListModel.StartDateTime = getEventResultModel?.start_datetime ?? DateTime.Now;
            eventListModel.EndDateTime = getEventResultModel?.end_datetime ?? DateTime.Now;
            eventListModel.EventDisplay = $"{eventListModel.EventName} {eventListModel.EventDate}";
            eventListModel.attendees = eventAttendance;

            if (getEventResultModel != null && getEventResultModel.attendance != null && (getEventResultModel.attendance.attendee_members != null || getEventResultModel.attendance.participant_members != null))
            {
                foreach (var attended in getEventResultModel.attendance.attendee_members)
                {
                    if (eventAttendance.Any(a => a.member_number == attended.member_number))
                    {
                        eventAttendance.Where(a => a.member_number == attended.member_number).Single().attended = true;
                        eventAttendance.Where(a => a.member_number == attended.member_number).Single().pal = "Y";
                    }
                    else
                    {
                        // Load out of unit members
                        eventAttendance.Add(new EventAttendance
                        {
                            id = attended.id,
                            first_name = attended.first_name,
                            last_name = attended.last_name,
                            member_number = attended.member_number,
                            patrol_name = "",
                            isAdultMember = false,
                            attended = true,
                            pal = "Y"
                        });
                    }
                }
                foreach (var attended in getEventResultModel.attendance.participant_members)
                {
                    if (eventAttendance.Any(a => a.member_number == attended.member_number))
                    {
                        eventAttendance.Where(a => a.member_number == attended.member_number).Single().pal = "P";
                    }
                    else
                    {
                        // for older events participant_members seems to be used, not attendee_members
                        // Load out of unit members
                        eventAttendance.Add(new EventAttendance
                        {
                            id = attended.id,
                            first_name = attended.first_name,
                            last_name = attended.last_name,
                            member_number = attended.member_number,
                            patrol_name = "",
                            isAdultMember = false,
                            attended = true,
                            pal = "P"
                        });
                    }
                }
                foreach (var assisted in getEventResultModel.attendance.assistant_members)
                {
                    if (eventAttendance.Any(a => a.member_number == assisted.member_number))
                    {
                        eventAttendance.Where(a => a.member_number == assisted.member_number).Single().pal = "A";
                    }
                }
                foreach (var lead in getEventResultModel.attendance.leader_members)
                {
                    if (eventAttendance.Any(a => a.member_number == lead.member_number))
                    {
                        eventAttendance.Where(a => a.member_number == lead.member_number).Single().pal = "L";
                    }
                }
                eventListModel.attendees = eventAttendance;
                return eventListModel;
            }

            return eventListModel;
        }

        public async Task<AttendanceReportModel> GenerateAttendanceReportData(DateTime fromDate, DateTime toDate, string selectedCalendar, string groupCalendar)
        {
            var attendanceReport = new AttendanceReportModel();
            var attendanceReportItems = new List<AttendanceReportItemModel>();
            await SetCalendar(selectedCalendar);

            var programEvents = await GetEventsForDates(fromDate, toDate);
            await ResetCalendar();
            if (!string.IsNullOrEmpty(groupCalendar))
            {
                await SetCalendar(groupCalendar);
                var groupProgramEvents = await GetEventsForDates(fromDate, toDate);
                await ResetCalendar();
                programEvents = programEvents.Concat(groupProgramEvents).ToList();
            }
            foreach (var programEvent in programEvents.OrderBy(pe => pe.StartDateTime))
            {
                var eventListModel = await GetAttendanceForEvent(programEvent.Id);
                programEvent.attendees = eventListModel.attendees;
                foreach (var attendee in programEvent.attendees)
                {
                    attendanceReportItems.Add(new AttendanceReportItemModel
                    {
                        MemberId = attendee.id,
                        MemberName = $"{attendee.first_name} {attendee.last_name}",
                        MemberFirstName = attendee.first_name,
                        MemberLastName = attendee.last_name,
                        EventName = programEvent.EventName,
                        EventChallengeArea = programEvent.ChallengeArea,
                        EventStartDate = programEvent.StartDateTime,
                        EventNameDisplay = $"{programEvent.EventName} {programEvent.EventDate}",
                        Attended = attendee.attended ? 1 : 0,
                        IsAdultMember = attendee.isAdultMember ? 1 : 0,
                        EventStatus = programEvent.EventStatus,
                        Pal = attendee.pal
                    });
                }
            }
            attendanceReport.attendanceReportItems = attendanceReportItems;

            var memberSummaries = new List<AttendanceReportMemberSummaryModel>();
            var attendanceReportItemsGroupedByMember = attendanceReportItems.GroupBy(a => a.MemberId);
            var totalEvents = attendanceReportItems.DistinctBy(i => i.EventNameDisplay).Where(ma => ma.EventStartDate <= DateTime.Now).Count();
            foreach (var memberAttendance in attendanceReportItemsGroupedByMember)
            {
                var attendedCount = memberAttendance.Where(ma => ma.EventStartDate <= DateTime.Now).Sum(ma => ma.Attended);
                memberSummaries.Add(new AttendanceReportMemberSummaryModel
                {
                    MemberId = memberAttendance.Key,
                    MemberName = memberAttendance.FirstOrDefault()?.MemberName ?? "",
                    IsAdultMember = memberAttendance.FirstOrDefault()?.IsAdultMember ?? 0,
                    AttendanceCount = attendedCount,
                    TotalEvents = totalEvents
                });
            }

            foreach (var attendanceItem in attendanceReportItems)
            {
                var attendanceCount = memberSummaries.Where(ms => ms.MemberId == attendanceItem.MemberId).FirstOrDefault()?.AttendanceCount ?? 0;
                var attendanceRate = totalEvents == 0 ? 0 : (decimal)attendanceCount / totalEvents * 100m;
                attendanceItem.MemberNameAndRate = $"{attendanceItem.MemberName} ({Math.Round(attendanceRate, 0)}%)";
            }

            var challengeAreaSummaries = new List<AttendanceReportChallengeAreaSummaryModel>();
            var programEventsGroupedByChallengeArea = programEvents.GroupBy(a => a.ChallengeArea);
            foreach (var challengeArea in programEventsGroupedByChallengeArea)
            {
                challengeAreaSummaries.Add(new AttendanceReportChallengeAreaSummaryModel
                {
                    ChallengeArea = challengeArea.Key,
                    EventCount = challengeArea.Count(),
                    TotalEvents = programEvents.Count()
                });
            }
            attendanceReport.attendanceReportChallengeAreaSummaries = challengeAreaSummaries;

            return attendanceReport;
        }


    }
}
