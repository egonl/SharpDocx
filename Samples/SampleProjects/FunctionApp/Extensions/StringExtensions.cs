namespace FunctionApp.Extensions;

public static class StringExtensions
{
    public static int ConvertToIntOrDefault(this string? input, int defaultValue)
    {
        if (string.IsNullOrWhiteSpace(input) || !int.TryParse(input, out var result))
            result = defaultValue;
        return result;
    }
}