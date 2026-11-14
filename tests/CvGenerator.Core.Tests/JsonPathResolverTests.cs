using System.Text.Json;
using CvGenerator.Core.Services;

namespace CvGenerator.Core.Tests;

public class JsonPathResolverTests
{
    private readonly JsonPathResolver _sut = new();

    [Fact]
    public void Resolves_nested_and_simple_paths()
    {
        using var doc = JsonDocument.Parse("""{"name":"Ada","personalInfo":{"email":"a@b.c"},"n":42}""");
        Assert.Equal("Ada", _sut.GetString(doc.RootElement, "name"));
        Assert.Equal("a@b.c", _sut.GetString(doc.RootElement, "personalInfo.email"));
        Assert.Equal("42", _sut.GetString(doc.RootElement, "n"));
        Assert.Null(_sut.GetString(doc.RootElement, "missing"));
    }

    [Fact]
    public void Coerces_object_to_single_item_array()
    {
        using var doc = JsonDocument.Parse("""{"sample_array":{"a":1,"b":2}}""");
        var items = _sut.GetArrayItems(doc.RootElement, "sample_array", out var coerced, out var missing);
        Assert.False(missing);
        Assert.True(coerced);
        Assert.Single(items);
        Assert.Equal(1, items[0].GetProperty("a").GetInt32());
    }
}
