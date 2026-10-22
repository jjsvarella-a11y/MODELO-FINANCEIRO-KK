using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using System.Text.Json;
using System.Diagnostics;
using var document=SpreadsheetDocument.Open(args[0],false);
var validator=new OpenXmlValidator(FileFormatVersions.Microsoft365){MaxNumberOfErrors=300};
var results=new List<object>();
foreach(var e in validator.Validate(document)) {
 var item=new {type=e.ErrorType.ToString(),id=e.Id,description=e.Description,part=e.Part?.Uri.ToString(),path=e.Path?.XPath,node=e.Node?.OuterXml?.Substring(0,Math.Min(600,e.Node.OuterXml.Length))};results.Add(item);Console.WriteLine(JsonSerializer.Serialize(item));
}
File.WriteAllText(args.Length>1?args[1]:"/tmp/sdk_results.json",JsonSerializer.Serialize(results,new JsonSerializerOptions{WriteIndented=true}));Console.WriteLine("ERRORS "+results.Count);

Environment.ExitCode=results.Count==0?0:1;
