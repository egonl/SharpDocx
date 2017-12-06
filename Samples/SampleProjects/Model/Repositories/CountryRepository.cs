using System;
using System.Collections.Generic;
using Model.Models;

namespace Model.Repositories
{
    public class CountryRepository
    {
        public static List<Country> GetCountries()
        {
            return new List<Country>
            {
                new Country
                {
                    Name = "Egypt",
                    Population = 96084500,
                    DateProclaimed = new DateTime(1953, 6, 18)
                },
                new Country
                {
                    Name = "Germany",
                    Population = 82349400,
                    DateProclaimed = new DateTime(1990, 10, 3)
                },
                new Country
                {
                    Name = "Ireland",
                    Population = 6378000,
                    DateProclaimed = new DateTime(1916, 4, 24)
                },
                new Country
                {
                    Name = "Israel",
                    Population = 8782260,
                    DateProclaimed = new DateTime(1948, 5, 14)
                },
                new Country
                {
                    Name = "Poland",
                    Population = 38634007,
                    DateProclaimed = new DateTime(1989, 9, 13)
                },
                new Country
                {
                    Name = "The Netherlands",
                    Population = 17170000,
                    DateProclaimed = new DateTime(1581, 7, 26)
                },
                new Country
                {
                    Name = "United States",
                    Population = 325365189,
                    DateProclaimed = new DateTime(1776, 7, 4)
                },
            };
        }
    }
}
