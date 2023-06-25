using ClosedXML.Excel;
using IxMilia.Stl;

namespace HeightScan;

public class HeightScan: IDisposable
{
    private readonly Plr40C plr40C = new Plr40C();
    private readonly Sn350 sn350 = new Sn350();

    private Measurement[][]? measurements;

    public float xMin { set; get; } = 21;
    public float xMax {set; get; } = 326;
    public float yMin {set; get; } = -10;
    public float yMax {set; get; } = 337;
    public float step {set; get; } = 10;
    
    public async Task<bool> ConnectDevices()
    {
        try
        {
            if (!await plr40C.Connect())
            {
                return false;
            }

            if (!await sn350.Connect())
            {
                return false;
            }

            return true;
        }
        catch (Exception e)
        {
            Console.WriteLine(e);
            return false;
        }
    }

    public async Task<bool> RunMeasurements()
    {
        try
        {
            if (!await sn350.AutoHome())
            {
                Console.WriteLine("AutoHome failed");
                return false;
            }

            int nrOfXSteps = (int)(((xMax - xMin) / step) - 1);  
            int nrOfYSteps = (int)(((yMax - yMin) / step) - 1);

            measurements = new Measurement[nrOfXSteps][];
            for (var i = 0; i < measurements.Length; i++)
            {
                measurements[i] = new Measurement[nrOfYSteps];
            }
        
            for (var y = 0; y < nrOfYSteps; y++)
            {
                for (var x = 0; x < nrOfXSteps; x++)
                {
                    var xPos = x * step + xMin;
                    var yPos = y * step + yMin;
                
                    if (!await sn350.GoTo(xPos, yPos, null, 1500, null))
                    {
                        Console.WriteLine("GoTo failed");
                        return false;
                    }
        
                    var distanceResult = await plr40C.MeasureDistance();
                    if (!distanceResult.success)
                    {
                        Console.WriteLine("Measure distance failed");
                        return false;
                    }

                    Console.WriteLine($"{x}, {y}, {xPos}, {yPos}, {distanceResult.distance}");

                    measurements[x][y] = new Measurement
                    {
                        x = xPos,
                        y = yPos,
                        z = (float)distanceResult.distance
                    };
                }
            }
        
            return true;
        }
        catch (Exception e)
        {
            Console.WriteLine(e);
            return false;
        }
    }

    public bool WriteExcel(string path)
    {
        try
        {
            if (measurements == null)
            {
                return false;
            }

            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("measurements");
            var row = worksheet.FirstRow();

            row.Cell(1).Value = "#x";
            row.Cell(2).Value = "#y";
            row.Cell(3).Value = "x";
            row.Cell(4).Value = "y";
            row.Cell(5).Value = "z";

            for (int x = 0; x < measurements.Length; x++)
            {
                for(int y = 0; y < measurements[x].Length; y++)
                {
                    row = row.RowBelow();
                    row.Cell(1).Value = x;
                    row.Cell(2).Value = y;
                    row.Cell(3).Value = measurements[x][y].x;
                    row.Cell(4).Value = measurements[x][y].y;
                    row.Cell(5).Value = measurements[x][y].z;
                }
            }
        
            workbook.SaveAs(path);

            return true;
        }
        catch (Exception e)
        {
            Console.WriteLine(e);
            return false;
        }
    }

    public bool WriteStl(string path)
    {
        try
        {
            if (measurements == null)
            {
                return false;
            }
            
            var stlFile = new StlFile();

            var xOffset = measurements[0][0].x;
            var yOffset = measurements[0][0].y;
            var zOffset = measurements[0][0].z;

            for (int x = 1; x < measurements.Length; x++)
            {
                for(int y = 1; y < measurements[x].Length; y++)
                {
                    var m00 = measurements[x - 1][y - 1];
                    var v00 = new StlVertex(m00.x - xOffset, m00.y - yOffset, zOffset - m00.z);
                    var m01 = measurements[x - 1][y];
                    var v01 = new StlVertex(m01.x - xOffset, m01.y - yOffset, zOffset - m01.z);
                    var m10 = measurements[x][y - 1];
                    var v10 = new StlVertex(m10.x - xOffset, m10.y - yOffset, zOffset - m10.z);
                    var m11 = measurements[x][y];
                    var v11 = new StlVertex(m11.x - xOffset, m11.y - yOffset, zOffset - m11.z);
                    var normal = new StlNormal(0, 0, 1);

                    var t1 = new StlTriangle(normal, v11, v01, v00);
                    stlFile.Triangles.Add(t1);
                    var t2 = new StlTriangle(normal, v11, v00, v10);
                    stlFile.Triangles.Add(t2);
                }
            }
            
            stlFile.Save(path);

            return true;
        }
        catch (Exception e)
        {
            Console.WriteLine(e);
            return false;
        }
    }

    public void Dispose()
    {
        plr40C.Dispose();
        sn350.Dispose();
    }
}