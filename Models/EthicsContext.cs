using System;
using System.Collections.Generic;
using Microsoft.EntityFrameworkCore;

namespace ExcelReport.Models;

public partial class EthicsContext : DbContext
{
    public EthicsContext()
    {
    }

    public EthicsContext(DbContextOptions<EthicsContext> options)
        : base(options)
    {
    }

    public virtual DbSet<Archive> Archives { get; set; }

    public virtual DbSet<Archivist> Archivists { get; set; }

    public virtual DbSet<Congress> Congresses { get; set; }

    public virtual DbSet<Doc> Docs { get; set; }

    public virtual DbSet<Inquiry> Inquiries { get; set; }

    public virtual DbSet<TestTable> TestTables { get; set; }

    //protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
    //{
    //    // Original scaffolded code removed or commented out
    //    // => optionsBuilder.UseSqlite("Data Source=C:\\HollingerBox\\EthicsXP_data.db");
    //}

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        modelBuilder.Entity<Archive>(entity =>
        {
            entity.HasKey(e => e.HascKey);

            entity.ToTable("Archive");

            entity.Property(e => e.HascKey).HasColumnName("HASC Key");
            entity.Property(e => e.ArchiveNo).HasColumnName("Archive No");
            entity.Property(e => e.BoxLabelProblem).HasColumnName("Box Label problem");
            entity.Property(e => e.BoxLabelWithoutCongress).HasColumnName("Box Label without congress");
            entity.Property(e => e.DocFound).HasColumnName("docFound");
            entity.Property(e => e.HollingerBoxKey).HasColumnName("Hollinger Box Key");
            entity.Property(e => e.Label1).HasColumnName("label1");
            entity.Property(e => e.Label2).HasColumnName("label2");
            entity.Property(e => e.Label3).HasColumnName("label3");
            entity.Property(e => e.Label4).HasColumnName("label4");

            entity.HasOne(d => d.CongressNavigation).WithMany(p => p.Archives)
                .HasForeignKey(d => d.Congress)
                .OnDelete(DeleteBehavior.ClientSetNull);

            entity.HasOne(d => d.SubcommitteeNavigation).WithMany(p => p.Archives)
                .HasForeignKey(d => d.Subcommittee)
                .OnDelete(DeleteBehavior.ClientSetNull);
        });

        modelBuilder.Entity<Archivist>(entity =>
        {
            entity.HasKey(e => e.Ric);

            entity.ToTable("Archivist");

            entity.Property(e => e.LoggedIn).HasColumnName("Logged in");
        });

        modelBuilder.Entity<Congress>(entity =>
        {
            entity.HasKey(e => e.CongressNo);

            entity.ToTable("Congress");

            entity.Property(e => e.CongressNo).ValueGeneratedNever();
            entity.Property(e => e.YearLabel).HasColumnName("Year Label");
        });

        modelBuilder.Entity<Doc>(entity =>
        {
            entity.HasKey(e => e.Key);

            entity.Property(e => e.DocDescrip).HasColumnName("Doc Descrip");
            entity.Property(e => e.HascKey).HasColumnName("HASC Key");
            entity.Property(e => e.UserId).HasColumnName("User ID");

            entity.HasOne(d => d.HascKeyNavigation).WithMany(p => p.Docs).HasForeignKey(d => d.HascKey);

            entity.HasOne(d => d.User).WithMany(p => p.Docs).HasForeignKey(d => d.UserId);
        });

        modelBuilder.Entity<Inquiry>(entity =>
        {
            entity.HasKey(e => e.Subcommittee);

            entity.ToTable("Inquiry");

            entity.Property(e => e.LongName).HasColumnName("Long Name");
        });

        modelBuilder.Entity<TestTable>(entity =>
        {
            entity.ToTable("TestTable");

            entity.Property(e => e.Id).ValueGeneratedNever();
        });

        OnModelCreatingPartial(modelBuilder);
    }

    partial void OnModelCreatingPartial(ModelBuilder modelBuilder);
}
