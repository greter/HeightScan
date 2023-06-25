using System.Net;
using Flurl;
using Flurl.Http;

namespace HeightScan;

//https://snapmaker.github.io/Documentation/gcode/G000-G001
public class Sn350: IDisposable
{
    private static string baseUrl = "http://192.168.10.83:8080/api/v1/";
    private static string token = "bc71c5fb-db72-4ff1-9897-2ebc0cd1d85d";
    private static bool connected = false;
    private readonly Task keepConnectionAlive = new Task(KeepConnectionAlive);
    public async Task<bool> Connect()
    {
        try
        {
            var result = await (baseUrl + "connect").PostUrlEncodedAsync(new {token = token});
            var status = await result.GetJsonAsync();

            token = status.token;
            connected = true;
            keepConnectionAlive.Start();
            
            return true;
        }
        catch (Exception e)
        {
            Console.WriteLine(e);
            return false;
        }
    }

    public async Task<bool> AutoHome()
    {
        try
        {
            await ExecCode("G53"); // Move in Machine Coordinates https://snapmaker.github.io/Documentation/gcode/G053
            await ExecCode("G21"); // Millimeter Units https://snapmaker.github.io/Documentation/gcode/G021
            await ExecCode("G28"); // Auto Home https://snapmaker.github.io/Documentation/gcode/G028
            await ExecCode("G90"); // Absolute Positioning https://snapmaker.github.io/Documentation/gcode/G090
            return true;
        }
        catch (Exception e)
        {
            Console.WriteLine(e);
            return false;
        }
    }

    public async Task<bool> GoTo(float? x, float? y, float? z, float? f, float? e)
    {
        try
        {
            // Linear Move https://snapmaker.github.io/Documentation/gcode/G000-G001
            var code = "G0";
            if (x != null)
            {
                code += $" X{x}";
            }
            if (y != null)
            {
                code += $" Y{y}";
            }
            if (z != null)
            {
                code += $" Z{z}";
            }
            if (f != null)
            {
                code += $" F{f}";
            }
            if (e != null)
            {
                code += $" E{e}";
            }

            await ExecCode(code);
            await ExecCode("M400"); // Finish Moves https://snapmaker.github.io/Documentation/gcode/M400

            return true;
        }
        catch (Exception exception)
        {
            Console.WriteLine(exception);
            return false;
        }
    }

    public void Dispose()
    {
        connected = false;
        keepConnectionAlive.Wait();
    }

    private async Task<IFlurlResponse> ExecCode(string code)
    {
        return await (baseUrl + "execute_code").PostUrlEncodedAsync(new {token = token, code = code});
    }
    private static async void KeepConnectionAlive()
    {
        while (connected)
        {
            var result = await (baseUrl + "status").SetQueryParam("token", token).GetJsonAsync();
            await Task.Delay(TimeSpan.FromSeconds(1));
        }
    }
}