namespace WebApi.Models
{
    public class CoDataInputSheet
    {
        public string CourseId { get; set; }
        public string AcademicPeriod { get; set; }
        public string CourseTitle { get; set; }
        public string StartDate { get; set; }
        public string EndDate { get; set; }
        public string MinEnrolled { get; set; }
        public string MaxEnrolled { get; set; }
        public object PriceGroupId { get; set; }
        public object CourseLevelid { get; set; }
        public object Reason { get; set; }
    }
}
