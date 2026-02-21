using BenchmarkDotNet.Attributes;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using SixLabors.ImageSharp.Memory;
using System;

namespace NPOI.Benchmarks;

[MemoryDiagnoser]
public class InsertPictureBenchmark
{
    private static readonly string[] lorem = @"Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut 
labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip 
ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat 
nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id 
est laborum.".Split(' ', '\r', '\n');

    private XSSFWorkbook workbook;
    private ISheet sheet1;

    [Params(500)]
    public int RowCount { get; set; }
    [Params(500)]
    public int PictureCount { get; set; }

    [IterationSetup]
    public void Setup()
    {
        workbook = new XSSFWorkbook();
        sheet1 = workbook.CreateSheet("Sheet1");

        for (var rowNum = 1; rowNum <= RowCount; rowNum++)
        {
            var row = sheet1.CreateRow(rowNum);
            row.CreateCell(0);
        }
    }

    [Benchmark]
    public void InsertPictures()
    {
        var drawing = sheet1.CreateDrawingPatriarch();
        var imageData = LoadPNGImage("dotnet.png", workbook);

        for(var rowNum = 1; rowNum <= PictureCount; rowNum++)
        {
            var anchor = new XSSFClientAnchor(0, 0, 0, 0, 0, rowNum, 1, rowNum+1);
            anchor.AnchorType = AnchorType.MoveDontResize;
            int imageId = workbook.AddPicture(imageData, PictureType.PNG);
            var pic= drawing.CreatePicture(anchor, imageId);
        }
        /*using(var file = File.Create("pictures.xlsx"))
        {
            workbook.Write(file);
        }*/
    }
    
    protected byte[] LoadPNGImage(string path, IWorkbook wb)
    {
        using(FileStream file = File.OpenRead(path))
        {
            byte[] buffer =new byte[file.Length];
            file.Read(buffer, 0, (int) file.Length);
            return buffer;
        }
    }
    [IterationCleanup]
    public void Cleanup()
    {
        workbook.Dispose();
    }
}