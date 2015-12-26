using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Actigraph.Parser
{
   public class SubjectData
    {
        public string ID { get; set; }
        public string Hospital { get; set; }
        public string Cottage { get; set; }
        public int Age { get; set; }
        public string Gender { get; set; }
        public DateTime Date { get; set; }

        public string DayofWeek { get; set; }
        public double Sedentary { get; set; }
        public double Light { get; set; }
        public double Lifestyle { get; set; }
        public double Moderate { get; set; }
        public double Vigorous { get; set; }
        public double VeryVigorous { get; set; }
        public double Axix1Cpm { get; set; }
        public double VectorMagnitudeCpm { get; set; }
        public long StepsCount { get; set; }
        public double Time { get; set; }
        public long FreedsonBouts{get; set; }
        public long SedentaryBouts { get; set; }

    }

   public class SubjectValidAverages
    {
        public string Id { get; set; }
        public double AvgSedentary { get; set; }
        public double AvgLight { get; set; }
        public double AvgLifestyle { get; set; }
        public double AvgModerate { get; set; }
        public double AvgTime { get; set; }
        public long ValidDays { get; set; }

    }

   public class SubjectRecords
    {
        public IEnumerable<SubjectData> SubjectData { get; set; }
        public IEnumerable<SubjectData> SubjectValidData { get; set; }
        public SubjectValidAverages SubjectValidAverages { get; set; }
        public string SubjectId { get; set; }
    }

    public class TableData
    {
        public string Name { get; set; }
        public string Value { get; set; }
    }

    public class SummaryTableData
    {
       
        public string Date { get; set; }

        public string Day_of_Week { get; set; }
        public double Sedentary { get; set; }
        public double Light { get; set; }
        public double LifeStyle { get; set; }
        public double Moderate { get; set; }
        public double Movements_Per_Minute { get; set; }
        public long Steps { get; set; }
        public double Wear_Time { get; set; }
       
    }

}
