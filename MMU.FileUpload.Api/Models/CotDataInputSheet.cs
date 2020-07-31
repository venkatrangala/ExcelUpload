namespace WebApi.Models
{
    public class CotDataInputSheet
    {
        public string CourseId { get; set; }
        public string CourseTitle { get; set; }
        public string StartDate { get; set; }
        public string EndDate { get; set; }
        public string MinEnrolled { get; set; }
        public string MaxEnrolled { get; set; }
        public object PriceGroupId { get; set; }
        public object CouseLevelid { get; set; }

    }
}
