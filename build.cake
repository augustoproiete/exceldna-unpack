#tool "nuget:?package=NuGet.CommandLine&version=6.13.1"
#tool "nuget:?package=7-Zip.CommandLine&version=18.1.0"

#addin "nuget:?package=Cake.MinVer&version=2.0.0"
#addin "nuget:?package=Cake.Args&version=1.0.1"
#addin "nuget:?package=Cake.7zip&version=4.1.0"

var target       = ArgumentOrDefault<string>("target") ?? "publish";
var buildVersion = MinVer(s => s.WithTagPrefix("v").WithDefaultPreReleasePhase("preview"));

Task("clean")
    .Does(() =>
{
    CleanDirectories("./artifact/**");
    CleanDirectories("./**/^{bin,obj}");
});

Task("restore")
    .IsDependentOn("clean")
    .Does(() =>
{
    DotNetCoreRestore("./exceldna-unpack.sln", new DotNetCoreRestoreSettings
    {
        LockedMode = true,
    });

    NuGetRestore("./test/ExcelDnaUnpack.Tests.ExcelAddIn/ExcelDnaUnpack.Tests.ExcelAddIn.csproj", new NuGetRestoreSettings
    {
        NoCache = true,
        NonInteractive = true,
        PackagesDirectory = MakeAbsolute(new DirectoryPath("./packages")),
    });
});

Task("build")
    .IsDependentOn("restore")
    .DoesForEach(new[] { "Debug", "Release" }, (configuration) =>
{
    MSBuild("./exceldna-unpack.sln", settings => settings
        .SetConfiguration(configuration)
        .UseToolVersion(MSBuildToolVersion.VS2019)
        .WithTarget("Rebuild")
        .SetVersion(buildVersion.Version)
        .SetFileVersion(buildVersion.FileVersion)
        .SetContinuousIntegrationBuild()
    );
});

Task("test")
    .IsDependentOn("build")
    .Does(() =>
{
    var settings = new DotNetCoreTestSettings
    {
        Configuration = "Release",
        NoRestore = true,
        NoBuild = true,
    };

    var projectFiles = GetFiles("./test/**/*.csproj");
    foreach (var file in projectFiles)
    {
        DotNetCoreTest(file.FullPath, settings);
    }
});

Task("pack")
    .IsDependentOn("test")
    .Does(() =>
{
    var releaseNotes = $"https://github.com/augustoproiete/exceldna-unpack/releases/tag/v{buildVersion.Version}";

    DotNetCorePack("./src/ExcelDnaUnpack/ExcelDnaUnpack.csproj", new DotNetCorePackSettings
    {
        Configuration = "Release",
        NoRestore = true,
        NoBuild = true,
        IncludeSymbols = true,
        IncludeSource = true,
        OutputDirectory = "./artifact/nuget",
        MSBuildSettings = new DotNetCoreMSBuildSettings
        {
            Version = buildVersion.Version,
            PackageReleaseNotes = releaseNotes,
        },
    });
});

Task("publish")
    .IsDependentOn("pack")
    .DoesForEach(new[] { "win7-x64", "win7-x86" }, (runtime) =>
{
    DotNetCorePublish("./src/ExcelDnaUnpack/ExcelDnaUnpack.csproj", new DotNetCorePublishSettings
    {
        Framework = "net5.0",
        Runtime = runtime,
        Configuration = "Release",
        SelfContained = true,
        PublishSingleFile = true,
        PublishTrimmed = true,
        OutputDirectory = $"./artifact/standalone/{runtime}",
        MSBuildSettings = new DotNetCoreMSBuildSettings()
            .SetVersion(buildVersion.Version)
            .WithProperty("IncludeNativeLibrariesForSelfExtract", "true")
    });

    SevenZip(new SevenZipSettings
    {
        Command = new AddCommand
        {
            Files = GetFiles($"./artifact/standalone/{runtime}/exceldna-unpack.*"),
            Archive = new FilePath($"./artifact/standalone/exceldna-unpack-{buildVersion.Version}-{runtime}.zip"),
        }
    });
});

Task("push")
    .IsDependentOn("publish")
    .Does(() =>
{
    var url =  EnvironmentVariable("NUGET_URL");
    if (string.IsNullOrWhiteSpace(url))
    {
        Information("No NuGet URL specified. Skipping publishing of NuGet packages");
        return;
    }

    var apiKey =  EnvironmentVariable("NUGET_API_KEY");
    if (string.IsNullOrWhiteSpace(apiKey))
    {
        Information("No NuGet API key specified. Skipping publishing of NuGet packages");
        return;
    }

    var nugetPushSettings = new DotNetCoreNuGetPushSettings
    {
        Source = url,
        ApiKey = apiKey,
    };

    foreach (var nugetPackageFile in GetFiles("./artifact/nuget/*.nupkg"))
    {
        DotNetCoreNuGetPush(nugetPackageFile.FullPath, nugetPushSettings);
    }
});

RunTarget(target);
