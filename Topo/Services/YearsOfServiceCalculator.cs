namespace Topo.Services
{
    // Years of service is tracked per section (from IntroToSection, not overall Scouting
    // membership), and only counts whole years actually completed - a member gets the "2 years"
    // badge once they pass their second anniversary in the section, not on their first day of it.
    public static class YearsOfServiceCalculator
    {
        public static int? CalculateCompletedYears(DateTime? introToSection, DateTime asOf)
        {
            if (!introToSection.HasValue)
                return null;

            var years = asOf.Year - introToSection.Value.Year;
            if (introToSection.Value.AddYears(years) > asOf)
                years--;

            return years < 0 ? 0 : years;
        }
    }
}
