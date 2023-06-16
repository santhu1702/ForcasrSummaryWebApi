﻿// <auto-generated> This file has been auto generated by EF Core Power Tools. </auto-generated>
#nullable disable
using System;
using System.Collections.Generic;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata;

namespace ForcasrSummaryWebApi.MetaData
{
    public partial class budgetForecastContext : DbContext
    {
        public budgetForecastContext()
        {
        }

        public budgetForecastContext(DbContextOptions<budgetForecastContext> options)
            : base(options)
        {
        }

        public virtual DbSet<Per9> Per9 { get; set; }
        public virtual DbSet<TblDimFactData> TblDimFactData { get; set; }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Per9>(entity =>
            {
                entity.HasNoKey();

                entity.ToTable("per9");

                entity.Property(e => e.Brand)
                    .IsRequired()
                    .HasMaxLength(200);

                entity.Property(e => e.Category)
                    .IsRequired()
                    .HasMaxLength(200);

                entity.Property(e => e.SalesType)
                    .IsRequired()
                    .HasMaxLength(20)
                    .IsUnicode(false);

                entity.Property(e => e.Source)
                    .IsRequired()
                    .HasMaxLength(200);

                entity.Property(e => e.Subcategory)
                    .IsRequired()
                    .HasMaxLength(200)
                    .HasColumnName("subcategory");

                entity.Property(e => e.Year)
                    .HasMaxLength(10)
                    .IsUnicode(false);

                entity.Property(e => e._1).HasColumnName("1");

                entity.Property(e => e._10).HasColumnName("10");

                entity.Property(e => e._11).HasColumnName("11");

                entity.Property(e => e._12).HasColumnName("12");

                entity.Property(e => e._13).HasColumnName("13");

                entity.Property(e => e._2).HasColumnName("2");

                entity.Property(e => e._3).HasColumnName("3");

                entity.Property(e => e._4).HasColumnName("4");

                entity.Property(e => e._5).HasColumnName("5");

                entity.Property(e => e._6).HasColumnName("6");

                entity.Property(e => e._7).HasColumnName("7");

                entity.Property(e => e._8).HasColumnName("8");

                entity.Property(e => e._9).HasColumnName("9");
            });

            modelBuilder.Entity<TblDimFactData>(entity =>
            {
                entity.HasKey(e => e.DimFactId)
                    .HasName("PK__tblDimFa__879BBDEA94C69DDA");

                entity.ToTable("tblDimFactData");

                entity.Property(e => e.DimFactId).HasColumnName("DimFactID");

                entity.Property(e => e.Column01).HasMaxLength(200);

                entity.Property(e => e.Column02).HasMaxLength(200);

                entity.Property(e => e.Column03).HasMaxLength(200);

                entity.Property(e => e.Column04).HasMaxLength(200);

                entity.Property(e => e.Column05).HasMaxLength(200);

                entity.Property(e => e.Column06).HasMaxLength(200);

                entity.Property(e => e.Column07).HasMaxLength(200);

                entity.Property(e => e.Column08).HasMaxLength(200);

                entity.Property(e => e.Column09).HasMaxLength(200);

                entity.Property(e => e.Column10).HasMaxLength(200);

                entity.Property(e => e.Column11).HasMaxLength(200);

                entity.Property(e => e.Column12).HasMaxLength(200);

                entity.Property(e => e.Column13).HasMaxLength(200);

                entity.Property(e => e.Column14).HasMaxLength(200);

                entity.Property(e => e.Column15).HasMaxLength(200);

                entity.Property(e => e.Column16).HasMaxLength(200);

                entity.Property(e => e.Column17).HasMaxLength(200);

                entity.Property(e => e.Column18).HasMaxLength(200);

                entity.Property(e => e.Column19).HasMaxLength(200);

                entity.Property(e => e.Column20).HasMaxLength(200);

                entity.Property(e => e.Column21).HasMaxLength(200);

                entity.Property(e => e.Column22).HasMaxLength(200);

                entity.Property(e => e.Column23).HasMaxLength(200);

                entity.Property(e => e.Column24).HasMaxLength(200);

                entity.Property(e => e.Column25).HasMaxLength(200);

                entity.Property(e => e.Column26).HasMaxLength(200);

                entity.Property(e => e.Column27).HasMaxLength(200);

                entity.Property(e => e.Column28).HasMaxLength(200);

                entity.Property(e => e.Column29).HasMaxLength(200);

                entity.Property(e => e.Column30).HasMaxLength(200);

                entity.Property(e => e.Column31).HasMaxLength(200);

                entity.Property(e => e.Column32).HasMaxLength(200);

                entity.Property(e => e.Column33).HasMaxLength(200);

                entity.Property(e => e.Column34).HasMaxLength(200);

                entity.Property(e => e.Column35).HasMaxLength(200);

                entity.Property(e => e.Column36).HasMaxLength(200);

                entity.Property(e => e.Column37).HasMaxLength(200);

                entity.Property(e => e.Column38).HasMaxLength(200);

                entity.Property(e => e.Column39).HasMaxLength(200);

                entity.Property(e => e.Column40).HasMaxLength(200);

                entity.Property(e => e.CreatedBy)
                    .IsRequired()
                    .HasMaxLength(100)
                    .HasDefaultValueSql("('SYSTEM')");

                entity.Property(e => e.CreatedDate)
                    .HasColumnType("datetime")
                    .HasDefaultValueSql("(getdate())");

                entity.Property(e => e.IsActive)
                    .IsRequired()
                    .HasDefaultValueSql("((1))");

                entity.Property(e => e.UpdatedBy)
                    .IsRequired()
                    .HasMaxLength(100)
                    .HasDefaultValueSql("('SYSTEM')");

                entity.Property(e => e.UpdatedDate)
                    .HasColumnType("datetime")
                    .HasDefaultValueSql("(getdate())");
            });

            OnModelCreatingGeneratedProcedures(modelBuilder);
            OnModelCreatingPartial(modelBuilder);
        }

        partial void OnModelCreatingPartial(ModelBuilder modelBuilder);
    }
}