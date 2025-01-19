using System;
using System.Collections.Generic;

namespace ExcelReport.Models;

public partial class Archive
{
    public string HascKey { get; set; } = null!;

    public string Subcommittee { get; set; } = null!;

    public int ArchiveNo { get; set; }

    public int Congress { get; set; }

    public int? Classified { get; set; }

    public string Status { get; set; } = null!;

    public string? HollingerBoxKey { get; set; }

    public int? Printed { get; set; }

    public string? Note { get; set; }

    public int? BoxLabelWithoutCongress { get; set; }

    public int? BoxLabelProblem { get; set; }

    public int? DocFound { get; set; }

    public string? Label1 { get; set; }

    public string? Label2 { get; set; }

    public string? Label3 { get; set; }

    public string? Label4 { get; set; }

    public virtual Congress CongressNavigation { get; set; } = null!;

    public virtual ICollection<Doc> Docs { get; set; } = new List<Doc>();

    public virtual Inquiry SubcommitteeNavigation { get; set; } = null!;
}
