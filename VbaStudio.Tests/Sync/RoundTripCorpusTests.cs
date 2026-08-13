using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Abstractions.TestingHelpers;
using System.Text;
using VbaStudio.Core.Model;
using VbaStudio.Core.Sync;
using VbaStudio.Tests.Fakes;
using Xunit;

namespace VbaStudio.Tests.Sync;

public class RoundTripCorpusTests
{
    // UserFormMain.frm is deliberately excluded: FakeVbaProjectAccess doesn't
    // simulate the VBE's real export blank-line bug or .frx binary handling.
    // That path is only verifiable against real Excel - see Task 11 Step 5.
    public static IEnumerable<object[]> Fixtures()
    {
        yield return new object[] { "ClassWithPredeclaredId.cls", "clsSingleton", ModuleKind.Class };
        yield return new object[] { "ClassWithDefaultMember.cls", "clsDefaultMember", ModuleKind.Class };
        yield return new object[] { "ClassWithEnumerator.cls", "clsEnumerableCollection", ModuleKind.Class };
        yield return new object[] { "ModuleWithGreekFrenchLiterals.bas", "modLiterals", ModuleKind.Standard };
        yield return new object[] { "SheetDocumentModule.cls", "Sheet1", ModuleKind.Document };
    }

    [Theory]
    [MemberData(nameof(Fixtures))]
    public void PullPushPull_IsByteIdentical_Twice(string fileName, string moduleName, ModuleKind kind)
    {
        var extension = Path.GetExtension(fileName);
        var encoding = kind.SourceEncoding();
        var sourcePath = Path.Combine(AppContext.BaseDirectory, "TestData", fileName);
        var code = File.ReadAllText(sourcePath, encoding);

        var fake = new FakeVbaProjectAccess();
        fake.Add(new VbaModule(moduleName, kind, code, extension));
        var fs = new MockFileSystem();
        var sync = new SyncEngine(fake, fs, "src");

        sync.Pull();
        var first = fs.File.ReadAllBytes(fs.Path.Combine("src", moduleName + extension));

        Assert.Equal(code, encoding.GetString(first));

        sync.Push();
        sync.Pull();
        var second = fs.File.ReadAllBytes(fs.Path.Combine("src", moduleName + extension));

        Assert.Equal(first, second);
    }

    [Fact]
    public void Pull_NonCp1252Character_ThrowsInsteadOfSilentlyCorrupting()
    {
        var fake = new FakeVbaProjectAccess();
        fake.Add(new VbaModule("modGreek", ModuleKind.Standard, "Attribute VB_Name = \"modGreek\"\r\nPublic Const Greeting As String = \"Καλημέρα\"", ".bas"));
        var fs = new MockFileSystem();
        var sync = new SyncEngine(fake, fs, "src");

        Assert.Throws<EncoderFallbackException>(() => sync.Pull());
    }
}
