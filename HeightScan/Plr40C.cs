using System.Data.HashFunction.CRC;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Devices.Bluetooth;
using Windows.Devices.Bluetooth.GenericAttributeProfile;
using Windows.Devices.Enumeration;

namespace HeightScan;

// https://learn.microsoft.com/en-us/windows/uwp/devices-sensors/gatt-client
// https://github.com/piannucci/pymtprotocol/blob/15040e296d8a8ad01bc4284d6919b473760c2509/glm-server.py#L226


public class Plr40C: IDisposable
{
    // chapter "3.1.4 Request Mode options and checksum calculation" in MT_connectivity_protocol_1_2_9.pdf 
    private readonly ICRC crc = CRCFactory.Instance.Create(new CRCConfig
    {
        Polynomial = 0xA6,
        InitialValue = 0xAA,
        HashSizeInBits = 8,
        ReflectIn = false,
        ReflectOut = false,
        XOrOut = 0x00
    });

    private BluetoothLEDevice? leDevice;
    private GattDeviceService? service;
    private GattCharacteristic? characteristic;
    private static Task? characteristicChanged = null;
    private static byte[]? characteristicValue = null;

    public async Task<bool> Connect()
    {

        var selector = BluetoothLEDevice.GetDeviceSelectorFromDeviceName("Bosch PLR40C x0025");

        Console.Write("Find devices");
        var deviceInformationCollection = await DeviceInformation.FindAllAsync(selector);
        Console.WriteLine($" -> done: # devices found {deviceInformationCollection.Count}");

        if (!deviceInformationCollection.Any())
        {
            Console.WriteLine("No device found");
            return false;
        }

        var deviceInformation = deviceInformationCollection[0];
        var deviceInformationName = deviceInformation.Name;
        var deviceInformationId = deviceInformation.Id;

        Console.Write($"Connect to device: {deviceInformationName} {deviceInformationId}");
        leDevice = await BluetoothLEDevice.FromIdAsync(deviceInformationId);
        var gattServicesResult = await leDevice.GetGattServicesAsync();
        if (gattServicesResult.Status != GattCommunicationStatus.Success)
        {
            Console.WriteLine(" -> failed");
            return false;
        }

        Console.WriteLine(" -> done");

        // Open chrome://bluetooth-internals/#devices in chrome browser to inspect devices
        // this will help to find the right UUID's
        Console.Write("Find service with UUID \"00005301-0000-0041-5253-534f46540000\"");
        service = gattServicesResult.Services
            .FirstOrDefault(s => s.Uuid == Guid.Parse("00005301-0000-0041-5253-534f46540000"));
        
        if (service == null)
        {
            Console.WriteLine(" -> service not found");
            return false;
        }

        Console.WriteLine(" -> done");

        Console.Write("Find characteristic with UUID \"00004301-0000-0041-5253-534f46540000\"");
        var characteristicsResult = await service.GetCharacteristicsAsync();
        if (characteristicsResult.Status != GattCommunicationStatus.Success)
        {
            Console.WriteLine(" -> failed");
            return false;
        }

        characteristic = characteristicsResult.Characteristics
            .FirstOrDefault(c => c.Uuid == Guid.Parse("00004301-0000-0041-5253-534f46540000"));

        if (characteristic == null)
        {
            Console.WriteLine(" -> characteristic not found");
            return false;
        }

        Console.WriteLine(" -> done");

        Console.Write("Register for value change");
        var characteristicConfigurationResult = await characteristic.WriteClientCharacteristicConfigurationDescriptorAsync(
            GattClientCharacteristicConfigurationDescriptorValue.Indicate);

        if (characteristicConfigurationResult != GattCommunicationStatus.Success)
        {
            Console.WriteLine(" -> failed");
            return false;
        }

        characteristic.ValueChanged += Characteristic_ValueChanged;

        Console.WriteLine(" -> done");
        return true;
    }

    public async Task<bool> TurnLaserOn()
    {
        Console.Write("Turn laser on");

        var result = await SendCommand(65, Array.Empty<byte>());
        if (result == null)
        {
            Console.WriteLine($" -> error while sending command");
            return false;
        }
        if (result[0] != 0x00)
        {
            Console.WriteLine($" -> error {result[0]}");
            return false;
        }

        Console.WriteLine(" -> done");
        return true;
    }

    public async Task<(bool success, double distance)> MeasureDistance()
    {
        Console.Write("Measure Distance");

        var data = new byte[1];
        data[0] = 0; // Front, 5Hz, Auto Adjust, Single -> see MT_connectivity_protocol_LRF_command_set_2_5_0.pdf
        
        var result = await SendCommand(64, data);
        if (result == null)
        {
            Console.WriteLine($" -> error while sending command");
            return (false, Double.NaN);
        }
        if (result[0] != 0x00)
        {
            Console.WriteLine($" -> error {result[0]}");
            return (false, Double.NaN);
        }

        Console.WriteLine(" -> done");

        var length = BitConverter.ToInt32(result, 2);
        
        return (true, length * 0.05); // distance is returned as multiple 50 µm -> * 0.05 converts it to milimeters
    }

    public void Dispose()
    {
        service?.Dispose();
        leDevice?.Dispose();
        characteristicChanged?.Dispose();
    }
    
    private async Task<byte[]?> SendCommand(byte command, byte[] data)
    {
        if (characteristic == null)
        {
            return null;
        }
        
        var frame = new byte[4 + data.Length];
        frame[0] = 0xC0; // define request frame format -> LONG
        frame[1] = command;
        frame[2] = Convert.ToByte(data.Length);
        data.CopyTo(frame, 3);
        frame[^1] = crc.ComputeHash(frame.Take(frame.Length - 1).ToArray()).Hash[0];
        
        characteristicChanged = new Task(() => { });
        var result = await characteristic.WriteValueWithResultAsync(frame.AsBuffer());

        if (result.Status != GattCommunicationStatus.Success)
        {
            return null;
        }

        await characteristicChanged; // wait for result

        return characteristicValue;
    }
    
    // callback invoked when result of send a command is available
    private void Characteristic_ValueChanged(GattCharacteristic sender,
        GattValueChangedEventArgs args)
    {
        characteristicValue = args.CharacteristicValue.ToArray();
        characteristicChanged?.Start(); // inform main thread that result is available
    }
}