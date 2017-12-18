using System.Collections.Generic;

namespace Model.Models
{
    public class DocumentViewModel
    {
        public string Title { get; set; }
        public string Date { get; set; }
        public List<Country> Countries { get; set; }
    }
}