using System;
using System.Collections.Generic;

namespace ExcelReport.Models;

public partial class Inquiry
{
    public string Subcommittee { get; set; } = null!;

    public string? LongName { get; set; }

    public string? Password { get; set; }

    public virtual ICollection<Archive> Archives { get; set; } = new List<Archive>();
}
