using VbaStudio.Core.Excel;
using Xunit;

namespace VbaStudio.Tests.Excel;

public class VbaSourceTextTests
{
    [Fact]
    public void RemoveHeaderAttributeLines_StripsVersionAndAttributeBlock()
    {
        var lines = new[]
        {
            "VERSION 1.0 CLASS",
            "BEGIN",
            "  MultiUse = -1  'True",
            "END",
            "Attribute VB_Name = \"frmMain\"",
            "Attribute VB_GlobalNameSpace = False",
            "Attribute VB_Creatable = False",
            "Attribute VB_PredeclaredId = True",
            "Attribute VB_Exposed = False",
            "Private Sub UserForm_Initialize()",
            "End Sub"
        };

        var result = VbaSourceText.RemoveHeaderAttributeLines(lines);

        Assert.Equal(new[] { "Private Sub UserForm_Initialize()", "End Sub" }, result);
    }

    [Fact]
    public void RemoveHeaderAttributeLines_StandardModuleSingleAttributeLine_StripsJustThatLine()
    {
        var lines = new[]
        {
            "Attribute VB_Name = \"modCalc\"",
            "Option Explicit",
            "",
            "Sub Foo()",
            "End Sub"
        };

        var result = VbaSourceText.RemoveHeaderAttributeLines(lines);

        Assert.Equal(new[] { "Option Explicit", "", "Sub Foo()", "End Sub" }, result);
    }

    [Fact]
    public void TrimExtraLeadingBlankLine_RemovesVbeInsertedBlankLine()
    {
        var original = new[] { "Private Sub UserForm_Initialize()", "End Sub" }; // legitEmptyLineCount = 0

        var exported = new[]
        {
            "VERSION 1.0 CLASS",
            "BEGIN",
            "END",
            "Attribute VB_Name = \"frmMain\"",
            "",              // VBE-inserted extra blank line (the bug)
            "Private Sub UserForm_Initialize()",
            "End Sub"
        };

        var result = VbaSourceText.TrimExtraLeadingBlankLine(original, exported);

        Assert.Equal(new[]
        {
            "VERSION 1.0 CLASS",
            "BEGIN",
            "END",
            "Attribute VB_Name = \"frmMain\"",
            "Private Sub UserForm_Initialize()",
            "End Sub"
        }, result);
    }

    [Fact]
    public void TrimExtraLeadingBlankLine_LegitimateBlankLine_NotRemoved()
    {
        var original = new[] { "", "Private Sub UserForm_Initialize()", "End Sub" }; // legitEmptyLineCount = 1

        var exported = new[]
        {
            "VERSION 1.0 CLASS",
            "BEGIN",
            "END",
            "Attribute VB_Name = \"frmMain\"",
            "",              // matches the legitimate blank line - must be kept
            "Private Sub UserForm_Initialize()",
            "End Sub"
        };

        var result = VbaSourceText.TrimExtraLeadingBlankLine(original, exported);

        Assert.Equal(exported, result);
    }
}
