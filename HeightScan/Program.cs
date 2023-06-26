using HeightScan;
using IxMilia.Stl;

using var heightScan = new HeightScan.HeightScan();

if (! await heightScan.ConnectDevices())
{
    Console.WriteLine("Connection to devices failed");
    return -1;
}

// all in millimeters
heightScan.xMin = 100;
heightScan.yMin = 100;
heightScan.xMax = 140;
heightScan.yMax = 140;
heightScan.step = 10;

if (!await heightScan.RunMeasurements())
{
    Console.WriteLine("Running measurements failed");
    return -1;
}

if (!heightScan.WriteExcel(@"HeightScan.xlsx"))
{
    Console.WriteLine("Writing Excel file failed");
    return -1;
}

if (!heightScan.WriteStl(@"HeightScan.stl"))
{
    Console.WriteLine("Writing STL file failed");
    return -1;
}

return 0;
