using System;
using System.Collections.Generic;

namespace ExcelReport.Models;

public partial class Archivist
{
    public string Ric { get; set; } = null!;

    public string? First { get; set; }

    public string? Last { get; set; }

    public int? LoggedIn { get; set; }

    public string? Password { get; set; }

    public virtual ICollection<Doc> Docs { get; set; } = new List<Doc>();
}
