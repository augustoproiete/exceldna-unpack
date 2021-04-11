#tool "nuget:?package=7-Zip.CommandLine&version=18.1.0"

#addin "nuget:?package=Cake.MinVer&version=1.0.0"
#addin "nuget:?package=Cake.Args&version=1.0.0"
#addin "nuget:?package=Cake.7zip&version=1.0.3"

var target       = ArgumentOrDefault<string>("target") ?? "publish";
var buildVersion = MinVer(s => s.WithTagPrefix("v").WithDefaultPreReleasePhase("preview"));

Task("clean")
    .Does(() =>
{
    CleanDirectories("./artifacts/**");
    CleanDirectories("./src/**/bin");
    CleanDirectories("./src/**/obj");
    CleanDirectories("./test/**/bin");
    CleanDirectories("./test/**/obj");
});

Task("restore")
    .IsDependentOn("clean")
    .Does(() =>
{
    DotNetCoreRestore("./exceldna-unpack.sln", new DotNetCoreRestoreSettings
    {
        LockedMode = true,
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
        .WithProperty("Version", buildVersion.Version)
        .WithProperty("FileVersion", buildVersion.FileVersion)
        .WithProperty("ContinuousIntegrationBuild", "true")
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
        OutputDirectory = "./artifacts/nuget",
        ArgumentCustomization = args =>
            args.AppendQuoted($"-p:Version={buildVersion.Version}")
                .AppendQuoted($"-p:PackageReleaseNotes={releaseNotes}")
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
        OutputDirectory = $"./artifacts/standalone/{runtime}",
        ArgumentCustomization = args =>
            args.AppendQuoted($"-p:Version={buildVersion.Version}")
                .Append($"-p:IncludeNativeLibrariesForSelfExtract=true")
    });

    SevenZip(new SevenZipSettings
    {
        Command = new AddCommand
        {
            Files = GetFiles($"./artifacts/standalone/{runtime}/exceldna-unpack.*"),
            Archive = new FilePath($"./artifacts/standalone/exceldna-unpack-{buildVersion.Version}-{runtime}.zip"),
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

    foreach (var nugetPackageFile in GetFiles("./artifacts/nuget/*.nupkg"))
    {
        DotNetCoreNuGetPush(nugetPackageFile.FullPath, nugetPushSettings);
    }
});

RunTarget(target);
