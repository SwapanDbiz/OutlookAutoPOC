using System.Text.Json;

namespace ZCO2.Utils;

public class TestData
{
    public string? BaseUrl { get; set; }
    public Dictionary<string, User>? Users { get; set; }
    public Email? Email { get; set; }
}

public class User
{
    public string? Username { get; set; }
    public string? Password { get; set; }
}

public class Email
{
    public string? Subject { get; set; }
    public string? Body { get; set; }
}

public class TestDataLoader
{
    private static TestData? _testData;

    public static TestData LoadTestData()
    {
        if (_testData == null)
        {
            string jsonPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "TestData", "testData.json");
            string jsonContent = File.ReadAllText(jsonPath);
            var options = new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            };
            _testData = JsonSerializer.Deserialize<TestData>(jsonContent, options)
                        ?? throw new InvalidOperationException("Unable to deserialize testData.json");

            if (_testData.Users == null || _testData.Email == null)
                throw new InvalidOperationException("testData.json is missing required sections 'users' or 'email'.");
        }
        return _testData;
    }
}