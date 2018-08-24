using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Formatting;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;

// http://www.intstrings.com/ramivemula/articles/file-upload-using-multipartformdatastreamprovider-in-asp-net-webapi/
// https://stackoverflow.com/questions/45364128/how-to-parse-form-data-using-azure-functions

namespace DocxValidatorV1
{
    public static class Validator
    {
        [FunctionName("Validate")]
        public static async Task<HttpResponseMessage> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = null)] HttpRequestMessage req, TraceWriter log)
        {
            log.Info("C# HTTP trigger function processed a request.");

            var filePath = "temp";
            var provider = new MultipartFormDataStreamProvider(filePath);
            await req.Content.ReadAsMultipartAsync(provider);

            var errors = new List<ValidationError>();
            foreach (var file in provider.FileData)
            {
                Console.WriteLine(Path.GetFileName(file.LocalFileName));
                errors = ValidateWordDocument(file.LocalFileName);
            }

            var response = new
            {
                errors
            };

            return req.CreateResponse(errors.Count == 0 ? HttpStatusCode.OK : HttpStatusCode.BadRequest, response,
                JsonMediaTypeFormatter.DefaultMediaType);
        }

        public static List<ValidationError> ValidateWordDocument(string filepath)
        {
            var errors = new List<ValidationError>();
            using (var wordprocessingDocument =
                WordprocessingDocument.Open(filepath, true))
            {
                var validator = new OpenXmlValidator();
                foreach (var error in
                    validator.Validate(wordprocessingDocument))
                    errors.Add(new ValidationError
                    {
                        Description = error.Description,
                        ErrorType = error.ErrorType.ToString(),
                        Node = error.Node.ToString(),
                        Path = error.Path.XPath,
                        Part = error.Part.Uri.ToString()
                    });

                wordprocessingDocument.Close();
                return errors;
            }
        }
    }
}