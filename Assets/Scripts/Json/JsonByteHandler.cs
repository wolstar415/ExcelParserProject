#if false

using System.IO;
using System.Text;
using System.IO.Compression;
using Newtonsoft.Json;
using UnityEngine;
#if UNITY_EDITOR
using UnityEditor;
#endif

public static class JsonByteHandler
{
    private static string dataPath = Path.Combine(Application.dataPath, "Resources");

    public static byte[] SerializeToBytes<T>(T data)
    {
        string json = JsonConvert.SerializeObject(data, Formatting.None);
        return Encoding.UTF8.GetBytes(json);
    }

    public static T DeserializeFromBytes<T>(byte[] bytes)
    {
        string json = Encoding.UTF8.GetString(bytes);
        return JsonConvert.DeserializeObject<T>(json);
    }

    public static byte[] Compress(byte[] data)
    {
        using (MemoryStream output = new MemoryStream())
        {
            using (GZipStream gzip = new GZipStream(output, CompressionMode.Compress))
            {
                gzip.Write(data, 0, data.Length);
            }
            return output.ToArray();
        }
    }

    public static byte[] Decompress(byte[] data)
    {
        using (MemoryStream input = new MemoryStream(data))
        using (GZipStream gzip = new GZipStream(input, CompressionMode.Decompress))
        using (MemoryStream output = new MemoryStream())
        {
            gzip.CopyTo(output);
            return output.ToArray();
        }
    }

    public static void SaveCompressedData<T>(T data, string fileName = "data.bytes")
    {
        if (!Directory.Exists(dataPath))
        {
            Directory.CreateDirectory(dataPath);
        }
        string filePath = Path.Combine(dataPath, fileName);
        byte[] jsonBytes = SerializeToBytes(data);
        byte[] compressedBytes = Compress(jsonBytes);
        File.WriteAllBytes(filePath, compressedBytes);

#if UNITY_EDITOR
        AssetDatabase.Refresh();
#endif
    }

    public static T LoadCompressedData<T>(string fileName = "data")
    {
        var _data = Resources.Load<TextAsset>(fileName);

        if (_data == null)
        {
            return default;
        }

        byte[] compressedBytes = _data.bytes;
        byte[] decompressedBytes = Decompress(compressedBytes);

        return DeserializeFromBytes<T>(decompressedBytes);
    }
}

#endif