using System;
using System.Collections.Generic;
using System.Text;

namespace MMU.Functions.Models
{
    public class CourseOfferingTemplate
    {
        public int Id { get; set; }
        //public string CourseId { get; set; }
        //public string CourseTitle { get; set; }
        public int MinEnrolled { get; set; }
        public int MaxEnrolled { get; set; }
        public int PriceGroupId { get; set; }
        public int CourseLevelId { get; set; }
        public int EnrollmentModeId { get; set; }
        //public DateTime StartDate{ get; set; }
        //public DateTime EndDate { get; set; }
    }
}
