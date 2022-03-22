using System;
using System.Collections.Generic;

namespace Model.Models
{
    public class MyViewModel
    {
        public string Title { get; set; }
        public string Date { get; set; }
        public List<Country> Countries { get; set; }
        public DateTime? AverageDateProclaimed { get; set; }
        public int? AveragePopulation { get; set; }
        public string Image64 { get; set; }
    }
}