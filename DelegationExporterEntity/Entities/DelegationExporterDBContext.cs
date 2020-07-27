using System;
using DelegationExporterEntity.Base;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata;

namespace DelegationExporterEntity.Entities
{
    public partial class DelegationExporterDBContext : DbContext
    {
        public DelegationExporterDBContext()
        {
        }

        public DelegationExporterDBContext(DbContextOptions<DelegationExporterDBContext> options)
            : base(options)
        {
        }

        public virtual DbSet<NecessaryFile> NecessaryFile { get; set; }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            if (!optionsBuilder.IsConfigured)
            {
#warning To protect potentially sensitive information in your connection string, you should move it out of source code. See http://go.microsoft.com/fwlink/?LinkId=723263 for guidance on storing connection strings.
                optionsBuilder.UseSqlServer(Constant.DB_CONN);
            }
        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<NecessaryFile>(entity =>
            {
                entity.Property(e => e.FileName)
                    .IsRequired()
                    .HasMaxLength(256);

                entity.Property(e => e.UpdateTime).HasColumnType("datetime");
            });

            OnModelCreatingPartial(modelBuilder);
        }

        partial void OnModelCreatingPartial(ModelBuilder modelBuilder);
    }
}
