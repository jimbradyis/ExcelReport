using System;
using System.Collections.Generic;

namespace ExcelReport.Models;

public partial class Congress
{
    public int CongressNo { get; set; }

    public string? Years { get; set; }

    public int? Default { get; set; }

    public string? YearLabel { get; set; }

    public string? Committee { get; set; }

    public virtual ICollection<Archive> Archives { get; set; } = new List<Archive>();
}
