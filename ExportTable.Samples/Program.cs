using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Diagnostics;
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
        [DisplayFormat(ConvertEmptyStringToNull = true, NullDisplayText = "<Not provided>")]
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
                //Email = "jane@doe.com",
                Email = "",
                //EmailConfirmed = true,
                FirstName = "Jane",
                LastName = "Doe",
                DateOfBirth = new DateTime(1992, 1, 31)
            });

            IEnumerable<Customer> customers2 = new List<Customer>();
            customers.ToTable(showHeader: true)
                .AddColumn(customer => customer.Id)
                .AddColumn(customer => customer.FirstName)
                .AddColumn(customer => customer.LastName)
                .AddColumn(customer => customer.Email)
                .AddColumn(customer => customer.EmailConfirmed)
                .AddColumn(customer => customer.DateOfBirth, format: "mm/yyyy")
                .GenerateSpreadsheet("customers.xlsx");


            IList<ValidationErrorInfo> errors;
            if (OpenXmlUtilities.ValidateSpreadsheet("customers.xlsx", out errors))
            {
                Process.Start("customers.xlsx");
            }
            else
            {
                foreach (var validationErrorInfo in errors)
                {
                    Console.WriteLine(validationErrorInfo.Description);
                }
            }

            Console.WriteLine(customers.ToTable(showHeader: false)
                .AddColumn(customer => customer.Id)
                .AddColumn(customer => customer.FirstName)
                .AddColumn(customer => customer.DateOfBirth, culture: new CultureInfo("en-US"))
                .AddColumn(customer => customer.DateOfBirth, culture: new CultureInfo("fr-FR"))
                .AddColumn(customer => customer.DateOfBirth, format: "{0:MM/yyyy}")
                .GenerateCsv());

            customers.ToTable(showHeader: true)
                .AddColumn(customer => customer.Id)
                .AddColumn(title: "Full Name", select: customer => customer.FirstName + " " + customer.LastName)
                .AddColumn(customer => customer.FirstName, format: "a: {0}", select: customer => customer.FirstName.Count(c => c == 'a'))
                .GenerateHtmlTable("customers.html");
        }
    }
}

