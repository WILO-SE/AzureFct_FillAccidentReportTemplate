using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Syncfusion.Presentation;
using System.Net.Http.Headers;
using System.Net.Http;
using System.Net;

namespace AzureFct_FillAccidentReportTemplate
{
    public static class Function1
    {
        [FunctionName("Function1")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequestMessage req, ILogger log)
        {            
            log.LogInformation("AzureFct_FillAccidentReportTemplate received a request");

            // [C# Code]
            //Gets the input PowerPoint stream from client request
            Stream stream = req.Content.ReadAsStreamAsync().Result;
            //Open the existing PowerPoint file
            using (IPresentation pptxDoc = Presentation.Open(stream))
            {
                //Gets the first slide from the PowerPoint Presentation
                ISlide slide = pptxDoc.Slides[0];
                //Gets the first shape of the slide
                IShape shape = (IShape)slide.Shapes[0];
                //Instance to hold paragraphs in textframe
                IParagraphs paragraphs = shape.TextBody.Paragraphs;
                //Adds paragraph to the textbody of shape
                IParagraph paragraph1 = paragraphs.Add();
                //Adds a TextPart to the paragraph
                ITextPart textpart1 = paragraph1.AddTextPart();
                //Adds text to the TextPart
                textpart1.Text = "Q&A SEGMENT";
                //Sets the color of the text as white
                textpart1.Font.Color = ColorObject.White;
                //Sets text as bold
                textpart1.Font.Bold = true;
                //Sets the font name of the text as Calibri
                textpart1.Font.FontName = "Calibri (Body)";
                //Sets the font size of the text
                textpart1.Font.FontSize = 25;
                //Adds paragraph to the textbody of shape
                paragraphs.Add();

                //Add a paragraph into paragraphs collection
                AddParagraph(paragraphs, "Q&A segment will be at the end of the webinar.");

                //Add a paragraph into paragraphs collection
                AddParagraph(paragraphs, "Please enter your questions in the Questions window.");

                //Add a paragraph into paragraphs collection
                AddParagraph(paragraphs, "A recording of the webinar will be available within a week.");

                //Create memory stream to save the output PowerPoint file
                MemoryStream memoryStream = new MemoryStream();
                //Save the PowerPoint into memory stream
                pptxDoc.Save(memoryStream);

                //Reset the memory stream position
                memoryStream.Position = 0;
                //Create the response to return
                HttpResponseMessage response = new HttpResponseMessage(HttpStatusCode.OK);
                //Set the PowerPoint saved stream as content of response
                response.Content = new ByteArrayContent(memoryStream.ToArray());
                //Set the contentDisposition as attachment
                response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
                {
                    FileName = "Sample.pptx"
                };
                //Set the content type as PPTX format mime type
                response.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.presentationml.presentation");
                //Return the response with output PowerPoint stream
                return response;
            }
        }

        private static void AddParagraph(IParagraphs paragraphs, string text)
        {
            //Adds paragraph to the textbody of shape
            IParagraph paragraph = paragraphs.Add();
            //Sets the list type of the paragraph as bulleted
            paragraph.ListFormat.Type = ListType.Bulleted;
            //Sets the list level as 1
            paragraph.IndentLevelNumber = 1;
            // Sets the hanging value
            paragraph.FirstLineIndent = -20;
            //Adds a TextPart to the paragraph
            ITextPart textpart = paragraph.AddTextPart();
            //Adds text to the TextPart
            textpart.Text = text;
            //Sets the color of the text as white
            textpart.Font.Color = ColorObject.White;
            //Sets the font name of the text as Calibri
            textpart.Font.FontName = "Calibri (Body)";
            //Sets the font size of the text
            textpart.Font.FontSize = 20;
            //Adds paragraph to the textbody of shape
            paragraphs.Add();
        }
    }
}
