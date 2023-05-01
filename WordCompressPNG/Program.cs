// See https://aka.ms/new-console-template for more information
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using static System.Net.Mime.MediaTypeNames;
using CloudConvert.API;
using CloudConvert.API.Models.ExportOperations;
using CloudConvert.API.Models.ImportOperations;
using CloudConvert.API.Models.JobModels;
using CloudConvert.API.Models.TaskOperations;
using System.Net;
using System.Configuration;
using CommandLine;
using WordCompressPNG;
using System.Diagnostics;
using CloudConvert.API.Models.TaskModels;

var commandLineArgs = Environment.GetCommandLineArgs();
var parserResult = Parser.Default.ParseArguments<Options>(commandLineArgs);
if (parserResult.Errors.Any())
{
    return;
}
 var options = parserResult.Value;

string? apiKey = options.ApiKey ?? ConfigurationManager.AppSettings["ApiKey"];
if (string.IsNullOrEmpty(apiKey) || apiKey.StartsWith("***"))
{
    Console.ForegroundColor = ConsoleColor.Red;
    Console.WriteLine("ERROR: Missing ApiKey in config file and in command line.");
    Console.ForegroundColor = ConsoleColor.Yellow;
    Console.WriteLine("Use --apikey on command line or include key in .config file.");
    Console.WriteLine(@"The file .config must have this structure:
<configuration>
	<appSettings>
		<add key=""ApiKey"" value=""*** Insert CloudConvert ApiKey here ***""/>
	</appSettings>
</configuration>");
    Console.ResetColor();
    return;
}

string? outputWordPath = options.Overwrite ? options.InputPath : options.OutputPath;
if (string.IsNullOrEmpty(options.InputPath))
{
    Console.ForegroundColor = ConsoleColor.Red;
    Console.WriteLine("ERROR: Missing input file.");
    Console.ResetColor();
    return;
}
if (string.IsNullOrEmpty(outputWordPath))
{
    Console.ForegroundColor = ConsoleColor.Red;
    Console.WriteLine("ERROR: Missing output file.");
    Console.ResetColor();
    return;
}

if (!options.Overwrite)
{
    File.Copy(options.InputPath, outputWordPath, true);
}
string filename = outputWordPath;

if (string.IsNullOrEmpty(apiKey) || apiKey.StartsWith("***"))
{
    Console.ForegroundColor = ConsoleColor.Red;
    Console.WriteLine("ERROR: Missing ApiKey in config file.");
    Console.ForegroundColor = ConsoleColor.Yellow;
    Console.WriteLine(@"The file .config must have this content:
<configuration>
<appSettings>
	<add key=""ApiKey"" value=""*** Insert CloudConvert ApiKey here ***""/>
</appSettings>
</configuration>");
    Console.ResetColor();
    return;
}
var _cloudConvert = new CloudConvertAPI(apiKey, options.Sandbox);

long totalOriginalPngSize = 0;
long totalCompressedPngSize = 0;

Console.WriteLine($"Loading {options.InputPath} file...");
using (var document = WordprocessingDocument.Open(filename, true))
{
    if (document.MainDocumentPart == null) throw new Exception("Invalid DOCX file");

    var pngImages = document.MainDocumentPart.ImageParts.Where(p => p.Uri.ToString().EndsWith(".png"));
    long totalFiles = pngImages.Count();
    Console.WriteLine($"Converting {totalFiles} PNG files...");

    var compressedFiles = 0;
    foreach (var imagePart in pngImages)
    {
        long originalSize;
        long finalSize;

        // Convert one PNG file
        var job = await _cloudConvert.CreateJobAsync(new JobCreateRequest
        {
            Tasks = new
            {
                upload = new ImportUploadCreateRequest(),
                optimize = new OptimizeCreateRequest
                {
                    Input = "upload",
                    Input_Format = CloudConvert.API.Models.Enums.OptimizeInputFormat.png,
                    //             Profile = CloudConvert.API.Models.Enums.OptimizeProfile.
                },
                download = new ExportUrlCreateRequest
                {
                    Input = "optimize",
                    Archive_Multiple_Files = false
                }
            },
            Tag = System.IO.Path.GetFileName(options.InputPath)
        });
        
        // Create temporary file
        string tempFilename = Path.GetTempFileName();

        // Copy png in temp file
        using (var readStream = imagePart.GetStream(FileMode.Open, FileAccess.Read))
        using (var tempFileStream = File.Create(tempFilename))
        {
            readStream.CopyTo(tempFileStream);
            originalSize = readStream.Position;
            readStream.Seek(0, SeekOrigin.Begin);
        }

        // Upload
        var uploadTask = job.Data.Tasks.FirstOrDefault(t => t.Name == "upload");
        if (uploadTask == null) throw new Exception("Upload task not found");

        byte[] file = await File.ReadAllBytesAsync(tempFilename);
        // In case of Sandbox use, the file must be whitelisted providing MD5 and filename
        string fileName = options.Sandbox ? @"test.png" : Path.GetFileName(tempFilename);
        await _cloudConvert.UploadAsync(uploadTask.Result.Form.Url.ToString(), file, fileName, uploadTask.Result.Form.Parameters);

        // Convert ...

        // Download
        job = await _cloudConvert.WaitJobAsync(job.Data.Id); // Wait for job completion
        uploadTask = job.Data.Tasks.FirstOrDefault(t => t.Name == "upload");
        if (uploadTask.Status == CloudConvert.API.Models.Enums.TaskStatus.error)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"Upload task failed: {uploadTask.Message}");
            Console.ResetColor();
            return;
        }

        var optimizeTask = job.Data.Tasks.FirstOrDefault(t => t.Name == "optimize");
        if (optimizeTask == null) throw new Exception("Optimize task not found");
        if (optimizeTask.Status == CloudConvert.API.Models.Enums.TaskStatus.error)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"Optimization task failed: {optimizeTask.Message}");
            Console.ResetColor();
            return;
        }

        var downloadTask = job.Data.Tasks.FirstOrDefault(t => t.Name == "download");
        if (downloadTask == null) throw new Exception("Download task not found");

        var fileExport = downloadTask.Result.Files.FirstOrDefault();
        if (fileExport == null) throw new Exception("Downloaded file not found");

#pragma warning disable SYSLIB0014 // Type or member is obsolete
        using (var client = new WebClient()) client.DownloadFile(fileExport.Url, tempFilename);
#pragma warning restore SYSLIB0014 // Type or member is obsolete

            // Rewrite converted file
        byte[] imageBytes = File.ReadAllBytes(tempFilename);
        var writePngStream = imagePart.GetStream(FileMode.Open, FileAccess.ReadWrite);
        using (BinaryWriter writer = new BinaryWriter(writePngStream))
        {
            writer.Write(imageBytes);
            writePngStream.SetLength(writePngStream.Position);
            finalSize = writePngStream.Position;
            writer.Close();
        }

        compressedFiles++;
        totalOriginalPngSize += originalSize;
        totalCompressedPngSize += finalSize;

        double savingRatio = 1 - (double)finalSize / (double)originalSize;
        Console.WriteLine($"Converted file {compressedFiles}/{totalFiles} from {originalSize:#,#} to {finalSize:#,#} saving {savingRatio:0.00%}%");

        File.Delete(tempFilename);

        // Stop if over MaxFiles and MaxFiles is defined
        if (options.MaxFiles > 0 && compressedFiles >= options.MaxFiles) break;
    }

    double savingRatioTotal = 1 - (double)totalCompressedPngSize / (double)totalOriginalPngSize;
    Console.ForegroundColor = ConsoleColor.Green;
    Console.WriteLine($"Converted {compressedFiles} files from {totalOriginalPngSize:#,#} to {totalCompressedPngSize:#,#} saving {savingRatioTotal:0.00%}%");
    Console.WriteLine($"Saving {outputWordPath}...");
    Console.ResetColor();

    document.Save();
}

