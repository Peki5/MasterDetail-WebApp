﻿// <auto-generated> This file has been auto generated by EF Core Power Tools. </auto-generated>
#nullable disable
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using Microsoft.EntityFrameworkCore;

namespace RPPP_WebApp.Models;

[Table("vrstaZah")]
public partial class VrstaZah
{
    [Key]
    [Column("idVrste")]
    public int IdVrste { get; set; }

    [Required]
    [Column("imeZahtjeva", TypeName = "text")]
    public string ImeZahtjeva { get; set; }

    [InverseProperty("IdVrsteNavigation")]
    public virtual ICollection<Zahtjev> Zahtjev { get; set; } = new List<Zahtjev>();
}