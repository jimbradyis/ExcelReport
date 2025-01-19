using System;
using System.Collections.Generic;

namespace ExcelReport.Models;

public partial class Doc
{
    public int Key { get; set; }

    public string DocDescrip { get; set; } = null!;

    public string? Action { get; set; }

    public string? HascKey { get; set; }

    public string? UserId { get; set; }

    public virtual Archive? HascKeyNavigation { get; set; }

    public virtual Archivist? User { get; set; }
}
