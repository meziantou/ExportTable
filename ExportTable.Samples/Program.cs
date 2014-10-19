using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Validation;

namespace ExportTable.Samples
{
    public class Customer
    {
        [Display(Name = "Identifier")]
        public int Id { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public DateTime DateOfBirth { get; set; }
        public string Email { get; set; }
        public bool EmailConfirmed { get; set; }
    }

    class Program
    {
        static void Main(string[] args)
        {
            List<Customer> customers = new List<Customer>();
            customers.Add(new Customer()
            {
                Id = 1,
                Email = "john@doe.com",
                EmailConfirmed = true,
                FirstName = "John",
                LastName = "Doe",
                DateOfBirth = new DateTime(1990, 1, 31)
            });

            customers.Add(new Customer()
            {
                Id = 2,
                Email = "jane@doe.com",
                EmailConfirmed = true,
                FirstName = "Jane",
                LastName = "Doe",
                DateOfBirth = new DateTime(1992, 1, 31)
            });

            customers.ToTable(showHeader: true)
                .AddColumns(customer => customer.Id, dataType: typeof(string))
                .AddColumns(customer => customer.FirstName)
                .AddColumns(customer => customer.Email)
                .AddColumns(customer => customer.EmailConfirmed)
                .AddColumns(customer => customer.DateOfBirth, format: "mm/yyyy")
                .GenerateSpreadsheet("customers.xlsx");


            IList<ValidationErrorInfo> errors;
            if (OpenXmlUtilities.ValidateSpreadsheet("customers.xlsx", out errors))
            {
                //Process.Start("customers.xlsx");
            }
            else
            {
                foreach (var validationErrorInfo in errors)
                {
                    Console.WriteLine(validationErrorInfo.Description);
                }
            }

            Console.WriteLine(customers.ToTable(showHeader: false)
                .AddColumns(customer => customer.Id)
                .AddColumns(customer => customer.FirstName)
                .AddColumns(customer => customer.DateOfBirth, culture: new CultureInfo("en-US"))
                .AddColumns(customer => customer.DateOfBirth, culture: new CultureInfo("fr-FR"))
                .AddColumns(customer => customer.DateOfBirth, format: "{0:MM/yyyy}")
                .GenerateCsv());

            customers.ToTable(showHeader: true)
                .AddColumns(customer => customer.Id)
                .AddColumns(title: "Full Name", select: customer => customer.FirstName + " " + customer.LastName)
                .AddColumns(customer => customer.FirstName, format: "a: {0}", select: customer => customer.FirstName.Count(c => c == 'a'))
                .GenerateHtmlTable("customers.html");
        }
    }
}

