using HeightScan;
using IxMilia.Stl;

using var heightScan = new HeightScan.HeightScan();

if (! await heightScan.ConnectDevices())
{
    Console.WriteLine("Connection to devices failed");
    return -1;
}

heightScan.xMin = 131;
heightScan.yMin = 94;
heightScan.xMax = 194;
heightScan.yMax = 222;
heightScan.step = 1;

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
