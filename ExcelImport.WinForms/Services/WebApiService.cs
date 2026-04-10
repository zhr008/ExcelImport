using System.Net.Http.Json;
using System.Text.Json;
using ExcelImport.Core.Models;

namespace ExcelImport.Services;

public sealed class WebApiService
{
    private static readonly JsonSerializerOptions JsonOptions = new(JsonSerializerDefaults.Web);
    private readonly HttpClient _httpClient = new();

    public async Task<WebApiImportResponse> SendAsync(WebApiOptions options, WebApiPayload payload, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(options.BaseUrl))
        {
            throw new InvalidOperationException("WebApi 已启用，但未配置 BaseUrl。");
        }

        var baseUrl = options.BaseUrl.TrimEnd('/');
        var endpoint = string.IsNullOrWhiteSpace(options.Endpoint) ? "/api/excel-import" : options.Endpoint;
        var requestUri = $"{baseUrl}{endpoint}";

        using var response = await _httpClient.PostAsJsonAsync(requestUri, payload, cancellationToken);
        var apiResponse = await ReadResponseAsync(response, cancellationToken);

        if (!response.IsSuccessStatusCode)
        {
            var message = string.IsNullOrWhiteSpace(apiResponse?.Message)
                ? $"WebApi 请求失败，状态码: {(int)response.StatusCode} ({response.ReasonPhrase})"
                : $"WebApi 请求失败: {apiResponse.Message}";
            throw new InvalidOperationException(message);
        }

        if (apiResponse is null)
        {
            throw new InvalidOperationException("WebApi 返回内容为空，无法确认导入结果。");
        }

        if (!apiResponse.Success)
        {
            throw new InvalidOperationException($"WebApi 返回失败结果: {apiResponse.Message}");
        }

        return apiResponse;
    }

    private static async Task<WebApiImportResponse?> ReadResponseAsync(HttpResponseMessage response, CancellationToken cancellationToken)
    {
        try
        {
            return await response.Content.ReadFromJsonAsync<WebApiImportResponse>(JsonOptions, cancellationToken);
        }
        catch (NotSupportedException)
        {
            return null;
        }
        catch (JsonException)
        {
            return null;
        }
    }
}
