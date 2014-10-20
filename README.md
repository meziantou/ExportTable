Export Table
===========

.NET library to export a list of objects to Excel (xlsx), HTML table or CSV.

**Export to Excel**

    customers.ToTable(showHeader: true)
        .AddColumn(customer => customer.Id)
        .AddColumn(customer => customer.FirstName)
        .AddColumn(customer => customer.LastName)
        .AddColumn(customer => customer.DateOfBirth, format: "mm/yyyy")
        .GenerateSpreadsheet("customers.xlsx")

**Export to HTML**

    customers.ToTable(showHeader: true)
        .AddColumn(customer => customer.Id)
        .AddColumn(customer => customer.FirstName)
        .AddColumn(customer => customer.LastName)
        .GenerateHtmlTable("customers.html") // Generate file
        .GenerateHtmlTable()                 // return string
        .GenerateHtmlTable(textWriter)       // write to stream

**Export to CSV**

    customers.ToTable(showHeader: true)
        .AddColumn(customer => customer.Id)
        .AddColumn(customer => customer.FirstName)
        .AddColumn(customer => customer.LastName)
        .GenerateCsv("customers.csv") // Generate file
        .GenerateCsv()                // return string
        .GenerateCsv(textWriter)      // write to stream

**Bounded / Unbounded Columns**

    customers.ToTable(showHeader: true)
        // Bounded
        .AddColumn(customer => customer.Id, title: "Identifier")
        // Unbounded
        .AddColumn(title: "Full Name", select: customer => customer.FirstName + " " + customer.LastName)

**Others**

Bounded columns automatically find display information such as column title or format from DataAnnotations :

    public class Customer
    {
        [Display(Name = "Identifier")]
        public int Id { get; set; }
        public string FirstName { get; set; }
        [DisplayFormat(ConvertEmptyStringToNull = true, NullDisplayText = "<Not provided>")]
        public string LastName { get; set; }
        [DisplayFormat(DataFormatString = "{0:MM/yyyy}")]
        public DateTime DateOfBirth { get; set; }
        public string Email { get; set; }
        public bool EmailConfirmed { get; set; }
    }

**Dependencies**

- [OpenXml 2.5 SDK](https://www.nuget.org/packages/DocumentFormat.OpenXml/)
- [CodeFluent Runtime Client](http://www.softfluent.com/products/codefluent-runtime-client)
