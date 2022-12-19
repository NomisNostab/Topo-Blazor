namespace Topo.Model.Program
{
    public class GetEventResultModel
    {
        public string id { get; set; } = "";
        public string status { get; set; } = "";
        public string title { get; set; } = "";
        public string location { get; set; } = "";
        public Organiser organiser { get; set; } = new Organiser();
        public Organiser[] organisers { get; set; } = new Organiser[0];
        public string challenge_area { get; set; } = "";
        public DateTime start_datetime { get; set; }
        public DateTime end_datetime { get; set; }
        public Upload[] uploads { get; set; } = new Upload[0];
        public Attendance attendance { get; set; } = new Attendance();
        public Invitee[] invitees { get; set; } = new Invitee[0];
        public Review review { get; set; } = new Review();
        public string owner_type { get; set; } = "";
        public string owner_id { get; set; } = "";
        public Achievement_Pathway_Oas_Data achievement_pathway_oas_data { get; set; } = new Achievement_Pathway_Oas_Data();
        public Achievement_Pathway_Logbook_Data achievement_pathway_logbook_data { get; set; } = new Achievement_Pathway_Logbook_Data();
    }

    public class Organiser
    {
        public string id { get; set; } = "";
        public string first_name { get; set; } = "";
        public string last_name { get; set; } = "";
        public string member_number { get; set; } = "";
        public string patrol_name { get; set; } = "";
    }

    public class Attendance
    {
        public Attendance_Members[] leader_members { get; set; } = new Attendance_Members[0];
        public Attendance_Members[] assistant_members { get; set; } = new Attendance_Members[0];
        public Attendance_Members[] participant_members { get; set; } = new Attendance_Members[0];
        public Attendance_Members[] attendee_members { get; set; } = new Attendance_Members[0];
    }

    public class Attendance_Members
    {
        public string id { get; set; } = "";
        public string first_name { get; set; } = "";
        public string last_name { get; set; } = "";
        public string member_number { get; set; } = "";
        public string patrol_name { get; set; } = "";
    }

    public class Review
    {
        public string general_rating { get; set; } = "";
        public string[] general_tags { get; set; } = new string[0];
        public string[] scout_method_elements { get; set; } = new string[0];
        public string[] scout_spices_elements { get; set; } = new string[0];
    }

    public class Achievement_Pathway_Oas_Data
    {
        public string award_rule { get; set; } = "";
        public Verifier verifier { get; set; } = new Verifier();
        public object[] groups { get; set; } = new object[0];
    }

    public class Verifier
    {
        public string name { get; set; } = "";
        public string contact { get; set; } = "";
        public string type { get; set; } = "";
    }

    public class Achievement_Pathway_Logbook_Data
    {
        public float distance_travelled { get; set; }
        public float distance_walkabout { get; set; }
        public Achievement_Meta achievement_meta { get; set; } = new Achievement_Meta();
        public object[] categories { get; set; } = new object[0];
        public Details details { get; set; } = new Details();
        public string title { get; set; } = "";
    }

    public class Achievement_Meta
    {
        public string stream { get; set; } = "";
        public string branch { get; set; } = "";
    }

    public class Details
    {
        public string activity_time_length { get; set; } = "";
        public string activity_grade { get; set; } = "";
    }

    public class Upload
    {
        public string id { get; set; } = "";
        public string filename { get; set; } = "";
        public string bucket { get; set; } = "";
        public string key { get; set; } = "";
        public string url { get; set; } = "";
        public DateTime uploaded_on { get; set; }
    }

    public class Invitee
    {
        public string invitee_id { get; set; } = "";
        public string invitee_type { get; set; } = "";
        public string invitee_name { get; set; } = "";
        public string id { get; set; } = "";
    }

}
