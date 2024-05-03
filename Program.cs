using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using P = DocumentFormat.OpenXml.Presentation;
using C15 = DocumentFormat.OpenXml.Office2013.Drawing.Chart;
using C14 = DocumentFormat.OpenXml.Office2010.Drawing.Charts;
using System.Xml.Linq;

X();
void X()
{
    string pptxFilePath = "C:\\Testing\\Template_Creation\\new_test_file\\NewFolder\\Ppt.pptx";

    using (PresentationDocument presentationDocument = PresentationDocument.Open(pptxFilePath, true))
    {
        PresentationPart presentationPart = presentationDocument.PresentationPart;
        if (presentationPart != null)
        {
            foreach (SlidePart slidePart in presentationPart.SlideParts)
            {
                ChartPart chartPart = slidePart.ChartParts.FirstOrDefault();
                if (chartPart != null)
                {
                    // Load the XML of the chart
                    XDocument chartXml;
                    using (var stream = chartPart.GetStream())
                    {
                        chartXml = XDocument.Load(stream);
                    }

                    // Find all <c:ser> elements
                    var seriesElements = chartXml.Descendants().Where(e => e.Name.LocalName == "ser");

                    foreach (var seriesElement in seriesElements)
                    {
                        var valElement = seriesElement.Descendants().FirstOrDefault(e => e.Name.LocalName == "val");

                        if (valElement != null)
                        {
                            // Extract numeric values
                            var numRefElement = valElement.Descendants().FirstOrDefault(e => e.Name.LocalName == "numRef");
                            if (numRefElement != null)
                            {
                                // Get the <c:v> elements under <c:numRef> for numeric values
                                var vElements = numRefElement.Descendants().Where(e => e.Name.LocalName == "v").ToList();

                                // Find the <c:cat> element for this series
                                var catElement = seriesElement.Descendants().FirstOrDefault(e => e.Name.LocalName == "cat");
                                if (catElement != null)
                                {
                                    // Get the <c:v> elements under <c:cat> for category labels
                                    var catVElements = catElement.Descendants().Where(e => e.Name.LocalName == "v").ToList();

                                    // Find the index of the category based on the entered category name
                                    string categoryNameToDelete = "Category 4"; // For example, enter the category name to delete
                                    int categoryIndexToDelete = catVElements.FindIndex(e => e.Value == categoryNameToDelete);

                                    if (categoryIndexToDelete != -1)
                                    {
                                        // Print out the category label and its corresponding value before deletion
                                        Console.WriteLine($"Category: {catVElements[categoryIndexToDelete].Value}, Value: {vElements[categoryIndexToDelete].Value}");

                                        // Remove the corresponding value
                                        if (categoryIndexToDelete < vElements.Count)
                                        {
                                            vElements[categoryIndexToDelete].Remove();
                                        }

                                        // Print out confirmation after deletion
                                        Console.WriteLine("Value deleted.");
                                    }
                                    else
                                    {
                                        Console.WriteLine($"Category '{categoryNameToDelete}' not found.");
                                    }
                                }
                            }
                        }
                    }

                    // Save the modified XML back to the chart part
                    using (var memoryStream = new MemoryStream())
                    {
                        chartXml.Save(memoryStream);
                        memoryStream.Seek(0, SeekOrigin.Begin);
                        chartPart.FeedData(memoryStream);
                    }

                    // Print a message after saving the modified XML
                    Console.WriteLine("Chart XML saved after modification.");

                    // Break out of the loop after processing the first slide
                    break;
                }
            }
        }

    }
}

