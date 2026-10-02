using ExcelTools.Core.Internal;
using OfficeTools.Common.Errors;

namespace ExcelTools.Core.Tests;

public class A1AddressTests
{
    [Theory]
    [InlineData(1, "A")]
    [InlineData(26, "Z")]
    [InlineData(27, "AA")]
    [InlineData(52, "AZ")]
    [InlineData(53, "BA")]
    [InlineData(702, "ZZ")]
    [InlineData(703, "AAA")]
    [InlineData(16384, "XFD")]
    public void ColumnName_known_values(int number, string name)
    {
        Assert.Equal(name, A1Address.ColumnName(number));
        Assert.Equal(number, A1Address.ParseRange($"{name}1").FirstColumn);
    }

    [Fact]
    public void Every_column_round_trips()
    {
        for (var c = 1; c <= A1Address.MaxColumn; c++)
        {
            var range = A1Address.ParseRange($"{A1Address.ColumnName(c)}1");
            Assert.Equal(c, range.FirstColumn);
        }
    }

    [Theory]
    [InlineData(0)]
    [InlineData(16385)]
    public void ColumnName_out_of_range_throws(int number) =>
        Assert.ThrowsAny<ArgumentOutOfRangeException>(() => A1Address.ColumnName(number));

    [Fact]
    public void Single_cell()
    {
        var r = A1Address.ParseRange("B3");
        Assert.Equal(new RangeAddress(3, 2, 3, 2), r);
        Assert.Equal(1, r.CellCount);
        Assert.Equal("B3", r.ToString());
    }

    [Fact]
    public void Last_cell_of_the_sheet()
    {
        var r = A1Address.ParseRange("XFD1048576");
        Assert.Equal(A1Address.MaxRow, r.FirstRow);
        Assert.Equal(A1Address.MaxColumn, r.FirstColumn);
        Assert.Equal("XFD1048576", r.ToString());
    }

    [Fact]
    public void Rectangle()
    {
        var r = A1Address.ParseRange("A1:C10");
        Assert.Equal(new RangeAddress(1, 1, 10, 3), r);
        Assert.Equal(10, r.RowCount);
        Assert.Equal(3, r.ColumnCount);
        Assert.Equal(30, r.CellCount);
        Assert.Equal("A1:C10", r.ToString());
    }

    [Fact]
    public void Reversed_corners_are_normalized()
    {
        Assert.Equal(A1Address.ParseRange("A1:C10"), A1Address.ParseRange("C10:A1"));
        Assert.Equal(A1Address.ParseRange("A1:C10"), A1Address.ParseRange("A10:C1"));
    }

    [Fact]
    public void Whole_column()
    {
        var r = A1Address.ParseRange("A:A");
        Assert.Equal(RangeKind.WholeColumns, r.Kind);
        Assert.Equal(new RangeAddress(1, 1, A1Address.MaxRow, 1, RangeKind.WholeColumns), r);
        Assert.Equal("A:A", r.ToString());
    }

    [Fact]
    public void Several_whole_columns_reversed()
    {
        var r = A1Address.ParseRange("C:A");
        Assert.Equal(1, r.FirstColumn);
        Assert.Equal(3, r.LastColumn);
        Assert.Equal("A:C", r.ToString());
    }

    [Fact]
    public void Whole_row()
    {
        var r = A1Address.ParseRange("1:1");
        Assert.Equal(RangeKind.WholeRows, r.Kind);
        Assert.Equal(new RangeAddress(1, 1, 1, A1Address.MaxColumn, RangeKind.WholeRows), r);
        Assert.Equal("1:1", r.ToString());
    }

    [Fact]
    public void Several_whole_rows()
    {
        var r = A1Address.ParseRange("5:2");
        Assert.Equal("2:5", r.ToString());
        Assert.Equal(4, r.RowCount);
    }

    [Fact]
    public void Whole_sheet_extremes_do_not_overflow_cell_count()
    {
        var r = A1Address.ParseRange("A1:XFD1048576");
        Assert.Equal(17_179_869_184L, r.CellCount);
    }

    [Theory]
    [InlineData("$A$1", "A1")]
    [InlineData("$A1", "A1")]
    [InlineData("A$1", "A1")]
    [InlineData("$A$1:$C$10", "A1:C10")]
    [InlineData("$A:$C", "A:C")]
    [InlineData("$1:$3", "1:3")]
    [InlineData("a1:c10", "A1:C10")]
    [InlineData("  A1:C10  ", "A1:C10")]
    [InlineData("A01", "A1")]
    [InlineData("A1:A1", "A1")]
    public void Tolerated_forms_normalize(string input, string expected) =>
        Assert.Equal(expected, A1Address.ParseRange(input).ToString());

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("A")]
    [InlineData("1")]
    [InlineData("A0")]
    [InlineData("0:0")]
    [InlineData("XFE1")]
    [InlineData("AAAA1")]
    [InlineData("A1048577")]
    [InlineData("A99999999")]
    [InlineData("1A")]
    [InlineData("A1:")]
    [InlineData(":A1")]
    [InlineData(":")]
    [InlineData("A1:B")]
    [InlineData("A:1")]
    [InlineData("A:B2")]
    [InlineData("A1:B2:C3")]
    [InlineData("A1 B2")]
    [InlineData("A 1")]
    [InlineData("$")]
    [InlineData("A$")]
    [InlineData("$$A1")]
    [InlineData("A-1")]
    [InlineData("甲1")]
    [InlineData("Sheet1!A1")]
    [InlineData("'My Sheet'!A1:B2")]
    public void Invalid_ranges_throw_INVALID_RANGE_with_format_hint(string input)
    {
        var ex = Assert.Throws<OfficeToolException>(() => A1Address.ParseRange(input));
        Assert.Equal(ErrorCodes.InvalidRange, ex.Code);
        Assert.Contains("A1:C10", ex.Hint);
    }

    [Fact]
    public void Null_is_invalid()
    {
        Assert.False(A1Address.TryParseRange(null, out _));
        Assert.Throws<OfficeToolException>(() => A1Address.ParseRange(null!));
    }

    [Fact]
    public void TryParse_reports_success_and_failure()
    {
        Assert.True(A1Address.TryParseRange("B2:D4", out var r));
        Assert.Equal("B2:D4", r.ToString());
        Assert.False(A1Address.TryParseRange("nope", out _));
    }

    [Fact]
    public void ParseCell_accepts_only_single_cells()
    {
        Assert.Equal(new CellAddress(10, 3), A1Address.ParseCell("C10"));
        Assert.Equal(new CellAddress(1, 1), A1Address.ParseCell("A1:A1"));
        Assert.Throws<OfficeToolException>(() => A1Address.ParseCell("A1:B2"));
        Assert.Throws<OfficeToolException>(() => A1Address.ParseCell("A:A"));
    }

    [Fact]
    public void Sheet_name_error_explains_where_the_sheet_goes()
    {
        var ex = Assert.Throws<OfficeToolException>(() => A1Address.ParseRange("Sheet1!A1"));
        Assert.Contains("工作表", ex.Message);
        Assert.Contains("sheet 參數", ex.Hint);
    }
}
