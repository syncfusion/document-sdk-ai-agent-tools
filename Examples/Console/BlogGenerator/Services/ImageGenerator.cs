using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using OpenAI;
using OpenAI.Images;

namespace BlogGenerator.Services
{
    /// <summary>
    /// Generates images via the OpenAI Images API and returns them as Base64 strings.
    /// </summary>
    public class ImageGenerator
    {
        private readonly ImageClient _imageClient;
        private readonly HttpClient _httpClient;
        private const int MaxRetries = 2;

        public ImageGenerator(OpenAIClient openAiClient, string model = "gpt-image-1.5")
        {
            _imageClient = openAiClient.GetImageClient(model);
            _httpClient = new HttpClient();
        }

        /// <summary>
        /// Generates an image from <paramref name="prompt"/> and returns the
        /// PNG bytes as a Base64-encoded string, or <c>null</c> if generation fails.
        /// </summary>
        public async Task<string?> GenerateBase64Async(string prompt)
        {
            for (int attempt = 1; attempt <= MaxRetries; attempt++)
            {
                try
                {
                    Console.WriteLine($"    [Image] Generating image (attempt {attempt})...");

                    var options = new ImageGenerationOptions
                    {
                        Size = GeneratedImageSize.W1024xH1024,
                    };

                    GeneratedImage image = await _imageClient.GenerateImageAsync(prompt, options);

                    if (image.ImageUri is not null)
                    {
                        var bytes = await _httpClient.GetByteArrayAsync(image.ImageUri);
                        return Convert.ToBase64String(bytes);
                    }

                    if (image.ImageBytes is not null)
                    {
                        //using var ms = new MemoryStream();
                        //await image.ImageBytes.CopyToAsync(ms);
                        return Convert.ToBase64String(image.ImageBytes.ToArray());
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"    [Image] Attempt {attempt} failed: {ex.Message}");
                    if (attempt < MaxRetries) await Task.Delay(2000 * attempt);
                }
            }

            Console.WriteLine("    [Image] Skipping image – all attempts failed.");
            return null;
        }
    }
}
