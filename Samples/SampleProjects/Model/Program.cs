using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Model.Models;
using Model.Repositories;
using SharpDocx;

namespace Model
{
    internal class Program
    {
        private static readonly string BasePath =
            Path.GetDirectoryName(typeof(Program).Assembly.Location) + @"\..\..\..\..";

        private static void Main()
        {
            var viewPath = $"{BasePath}\\Views\\Model.cs.docx";
            var documentPath = $"{BasePath}\\Documents\\Model.docx";

            //var countries = new List<Country>();
            //var countries = new List<Country> { new Country { Name = "Porosika" } };
            var countries = CountryRepository.GetCountries();

            var model = new MyViewModel
            {
                Title = "Model Sample",
                Date = DateTime.Now.ToShortDateString(),
                Countries = countries,
                AveragePopulation = countries.Count > 0
                    ? (int)countries.Average(c => c.Population)
                    : (int?)null,
                AverageDateProclaimed = GetAverageDateProclaimed(countries)
            };

#if DEBUG
            Ide.Start(viewPath, documentPath, model);
#else
            var document = DocumentFactory.Create(viewPath, model);
            var startTime = DateTime.Now;

            for (int i = 0; (DateTime.Now - startTime).TotalSeconds < 1; ++i)
            {
                // Create a new model for each document.
                //model = new MyViewModel { Title = $"Model Sample Part {i+1}", Countries = new List<Country>()};
                //document.Generate($"{BasePath}\\Documents\\Model {i}.docx", model);

                // Reuse the initial model.
                model.Title = $"Model Sample Part {i}";
                document.Generate($"{BasePath}\\Documents\\Model {i}.docx");
            }
#endif
        }

        private static DateTime? GetAverageDateProclaimed(List<Country> countries)
        {
            var dates = countries
                .Where(c => c.DateProclaimed.HasValue)
                .Select(c => c.DateProclaimed.Value)
                .ToList();

            if (dates.Count == 0)
            {
                return null;
            }

            return new DateTime((long) dates
                .Aggregate<DateTime, double>(0, (current, date) => current + date.Ticks / dates.Count));
        }
    }
}