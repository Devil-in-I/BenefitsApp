using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;
using MongoDB.EntityFrameworkCore;
using System.Text.Json.Serialization;

namespace BenefitsApp.Core.Models;

[Collection("Products")]
public class Benefit
{
    [BsonId]
    public required ObjectId Id { get; set; }

    [BsonRepresentation(MongoDB.Bson.BsonType.Int32)]
    public required int Code { get; set; }

    public required string Name { get; set; }

    public decimal RetailPrice { get; set; }

    public decimal DealerPrice { get; set; }

    public decimal SpecialPrice { get; set; }

    public int? WarrantyPeriod { get; set; }

    public string? Note { get; set; }

    public required string CategoryName { get; set; }
}
