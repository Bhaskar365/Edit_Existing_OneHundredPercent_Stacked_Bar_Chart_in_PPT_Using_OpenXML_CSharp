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
        // original email code

        //// Access the presentation part and then the slide part
        //PresentationPart presentationPart = presentationDocument.PresentationPart;
        //SlidePart slidePart = presentationPart.SlideParts.FirstOrDefault();

        //if (slidePart != null)
        //{
        //    // Create a new chart part
        //    ChartPart chartPart = slidePart.AddNewPart<ChartPart>();
        //    GenerateChartPart1Content(chartPart);

        //    // Create a new chart and set its properties
        //    A.Chart chart = new A.Chart();
        //    //chart.Append(new AutoTitleDeleted() { Val = true });

        //    // Create a new plot area
        //    PlotArea plotArea = new PlotArea();
        //    Layout layout = new Layout();

        //    // Add a bar chart and configure its properties
        //    BarChart barChart = plotArea.AppendChild(new BarChart(
        //        new BarDirection() { Val = BarDirectionValues.Column }));

        //    // Add data to the chart
        //    BarChartSeries barChartSeries = barChart.AppendChild(new BarChartSeries(
        //        new C.Index() { Val = UInt32Value.FromUInt32(0) },
        //        new Order() { Val = UInt32Value.FromUInt32(0) },
        //        new SeriesText(new NumericValue() { Text = "Series 1" })));

        //    // Specify the data range
        //    string formula = "Sheet1!$A$1:$A$5";
        //    StringReference stringReference = new StringReference { Formula = new Formula(formula) };
        //    barChartSeries.AppendChild(stringReference);

        //    // Append the plot area to the chart
        //    chart.Append(plotArea);

        //    // Append the chart to the chart part
        //    chartPart.ChartSpace = new ChartSpace();
        //    chartPart.ChartSpace.Append(chart);

        //    // Save the changes
        //    chartPart.ChartSpace.Save();

        //PresentationPart presentationPart = presentationDocument.PresentationPart;
        //if (presentationPart != null)
        //{
        //    foreach (SlidePart slidePart in presentationPart.SlideParts)
        //    {
        //        ChartPart chartPart = slidePart.ChartParts.FirstOrDefault();
        //        if (chartPart != null)
        //        {
        //            BarChart barChart = chartPart.ChartSpace.Descendants<BarChart>().FirstOrDefault();

        //            if (barChart != null)
        //            {
        //                // Iterate through the series in the bar chart
        //                foreach (BarChartSeries series in barChart.Descendants<BarChartSeries>())
        //                {

        //                    StringValue seriesName = series.SeriesText.InnerText;
        //                    Console.WriteLine($"Series Name: {seriesName}");

        //                    //NumberReference numberReference = series;

        //                    // Access the values for each series

        //                    //NumberReference numberReference = series.Values.NumberReference;
        //                    //NumberingCache numberingCache = numberReference.Formula.InnerText;
        //                    //Console.WriteLine($"Values: {numberingCache}");

        //                    //// Access the categories (axis labels)
        //                    //CategoryAxisData categories = series.CategoryAxisData;
        //                    //StringReference stringReference = categories.StringReference;
        //                    //StringCache stringCache = stringReference.StringCache;
        //                    //foreach (StringPoint point in stringCache.StringPoint)
        //                    //{
        //                    //    Console.WriteLine($"Category: {point.InnerText}");
        //                    //}
        //                }
        //            }
        //            else
        //            {
        //                Console.WriteLine("No bar chart found in the chart part.");
        //            }
        //        }

        //load xml values with labels
        //PresentationPart presentationPart = presentationDocument.PresentationPart;
        //if (presentationPart != null)
        //{
        //    foreach (SlidePart slidePart in presentationPart.SlideParts)
        //    {
        //        ChartPart chartPart = slidePart.ChartParts.FirstOrDefault();
        //        if (chartPart != null)
        //        {
        //            // Load the XML of the chart
        //            XDocument chartXml = XDocument.Load(chartPart.GetStream());

        //            // Find all <c:ser> elements
        //            var seriesElements = chartXml.Descendants().Where(e => e.Name.LocalName == "ser");

        //            foreach (var seriesElement in seriesElements)
        //            {
        //                // Get the <c:val> elements under <c:ser> for numeric values
        //                var valElement = seriesElement.Descendants().FirstOrDefault(e => e.Name.LocalName == "val");

        //                if (valElement != null)
        //                {
        //                    // Extract numeric values
        //                    var numRefElement = valElement.Descendants().FirstOrDefault(e => e.Name.LocalName == "numRef");
        //                    if (numRefElement != null)
        //                    {
        //                        // Get the <c:v> elements under <c:numRef> for numeric values
        //                        var vElements = numRefElement.Descendants().Where(e => e.Name.LocalName == "v").ToList();

        //                        // Find the <c:cat> element for this series
        //                        var catElement = seriesElement.Descendants().FirstOrDefault(e => e.Name.LocalName == "cat");
        //                        if (catElement != null)
        //                        {
        //                            // Get the <c:v> elements under <c:cat> for category labels
        //                            var catVElements = catElement.Descendants().Where(e => e.Name.LocalName == "v").ToList();

        //                            // Specify the index of the category to be deleted
        //                            int categoryIndexToDelete = 2; // For example, delete the third category
        //                            if (categoryIndexToDelete < catVElements.Count)
        //                            {
        //                                // Remove the category label
        //                                catVElements[categoryIndexToDelete].Remove();

        //                                // Remove the corresponding value
        //                                vElements[categoryIndexToDelete].Remove();
        //                            }
        //                        }
        //                    }
        //                }
        //            }

        //            // Save the modified XML back to the chart part
        //            chartXml.Save(chartPart.GetStream(FileMode.Create, FileAccess.Write));
        //        }
        //    }
        //}

        //PresentationPart presentationPart = presentationDocument.PresentationPart;
        //if (presentationPart != null)
        //{
        //    foreach (SlidePart slidePart in presentationPart.SlideParts)
        //    {
        //        ChartPart chartPart = slidePart.ChartParts.FirstOrDefault();
        //        if (chartPart != null)
        //        {
        //            // Load the XML of the chart
        //            XDocument chartXml;
        //            using (var stream = chartPart.GetStream())
        //            {
        //                chartXml = XDocument.Load(stream);
        //            }

        //            // Find all <c:ser> elements
        //            var seriesElements = chartXml.Descendants().Where(e => e.Name.LocalName == "ser");

        //            foreach (var seriesElement in seriesElements)
        //            {
        //                // Get the <c:val> elements under <c:ser> for numeric values
        //                var valElement = seriesElement.Descendants().FirstOrDefault(e => e.Name.LocalName == "val");

        //                if (valElement != null)
        //                {
        //                    // Extract numeric values
        //                    var numRefElement = valElement.Descendants().FirstOrDefault(e => e.Name.LocalName == "numRef");
        //                    if (numRefElement != null)
        //                    {
        //                        // Get the <c:v> elements under <c:numRef> for numeric values
        //                        var vElements = numRefElement.Descendants().Where(e => e.Name.LocalName == "v").ToList();

        //                        // Find the <c:cat> element for this series
        //                        var catElement = seriesElement.Descendants().FirstOrDefault(e => e.Name.LocalName == "cat");
        //                        if (catElement != null)
        //                        {
        //                            // Get the <c:v> elements under <c:cat> for category labels
        //                            var catVElements = catElement.Descendants().Where(e => e.Name.LocalName == "v").ToList();

        //                            // Specify the index of the category to be deleted
        //                            int categoryIndexToDelete = 2; // For example, delete the third category
        //                            if (categoryIndexToDelete < catVElements.Count)
        //                            {
        //                                // Print out the category label and its corresponding value before deletion
        //                                Console.WriteLine($"Category: {catVElements[categoryIndexToDelete].Value}, Value: {vElements[categoryIndexToDelete].Value}");

        //                                // Remove the category label
        //                                catVElements[categoryIndexToDelete].Remove();

        //                                // Remove the corresponding value
        //                                vElements[categoryIndexToDelete].Remove();

        //                                // Print out confirmation after deletion
        //                                Console.WriteLine("Category and its value deleted.");
        //                            }
        //                        }
        //                    }
        //                }
        //            }

        //            // Save the modified XML back to the chart part
        //            using (var stream = chartPart.GetStream(FileMode.Create, FileAccess.Write))
        //            {
        //                chartXml.Save(stream);
        //            }

        //            // Print a message after saving the modified XML
        //            Console.WriteLine("Chart XML saved after modification.");

        //            // Break out of the loop after processing the first slide
        //            break;
        //        }
        //    }
        //}

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
                            // Get the <c:val> elements under <c:ser> for numeric values
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

                                        // Specify the index of the category to be deleted
                                        int categoryIndexToDelete = 2; // For example, delete the third category
                                        if (categoryIndexToDelete < catVElements.Count)
                                        {
                                            // Print out the category label and its corresponding value before deletion
                                            Console.WriteLine($"Category: {catVElements[categoryIndexToDelete].Value}, Value: {vElements[categoryIndexToDelete].Value}");

                                            // Remove the category label
                                            catVElements[categoryIndexToDelete].Remove();

                                            // Remove the corresponding value
                                            vElements[categoryIndexToDelete].Remove();

                                            // Print out confirmation after deletion
                                            Console.WriteLine("Category and its value deleted.");
                                        }
                                    }
                                }
                            }
                        }

                        // Save the modified XML back to the chart part
                        using (var stream = chartPart.GetStream(FileMode.Create, FileAccess.Write))
                        {
                            chartXml.Save(stream);
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

void GenerateChartPart1Content(ChartPart chartPart1)
{
    C.ChartSpace chartSpace1 = new C.ChartSpace();
    chartSpace1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
    chartSpace1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
    chartSpace1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
    chartSpace1.AddNamespaceDeclaration("c16r2", "http://schemas.microsoft.com/office/drawing/2015/06/chart");
    C.Date1904 date19041 = new C.Date1904() { Val = false };
    C.EditingLanguage editingLanguage1 = new C.EditingLanguage() { Val = "en-US" };
    C.RoundedCorners roundedCorners1 = new C.RoundedCorners() { Val = false };

    AlternateContent alternateContent1 = new AlternateContent();
    alternateContent1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");

    AlternateContentChoice alternateContentChoice1 = new AlternateContentChoice() { Requires = "c14" };
    alternateContentChoice1.AddNamespaceDeclaration("c14", "http://schemas.microsoft.com/office/drawing/2007/8/2/chart");
    C14.Style style1 = new C14.Style() { Val = 102 };

    alternateContentChoice1.Append(style1);

    AlternateContentFallback alternateContentFallback1 = new AlternateContentFallback();
    C.Style style2 = new C.Style() { Val = 2 };

    alternateContentFallback1.Append(style2);

    alternateContent1.Append(alternateContentChoice1);
    alternateContent1.Append(alternateContentFallback1);

    C.Chart chart1 = new C.Chart();
    C.AutoTitleDeleted autoTitleDeleted1 = new C.AutoTitleDeleted() { Val = true };

    C.PlotArea plotArea1 = new C.PlotArea();
    C.Layout layout1 = new C.Layout();

    C.BarChart barChart1 = new C.BarChart();
    C.BarDirection barDirection1 = new C.BarDirection() { Val = C.BarDirectionValues.Bar };
    C.BarGrouping barGrouping1 = new C.BarGrouping() { Val = C.BarGroupingValues.PercentStacked };
    C.VaryColors varyColors1 = new C.VaryColors() { Val = false };

    C.BarChartSeries barChartSeries1 = new C.BarChartSeries();
    C.Index index1 = new C.Index() { Val = (UInt32Value)0U };
    C.Order order1 = new C.Order() { Val = (UInt32Value)0U };

    C.SeriesText seriesText1 = new C.SeriesText();

    C.StringReference stringReference1 = new C.StringReference();
    C.Formula formula1 = new C.Formula();
    formula1.Text = "Sheet1!$B$1";

    C.StringCache stringCache1 = new C.StringCache();
    C.PointCount pointCount1 = new C.PointCount() { Val = (UInt32Value)1U };

    C.StringPoint stringPoint1 = new C.StringPoint() { Index = (UInt32Value)0U };
    C.NumericValue numericValue1 = new C.NumericValue();
    numericValue1.Text = "Series 1";

    stringPoint1.Append(numericValue1);

    stringCache1.Append(pointCount1);
    stringCache1.Append(stringPoint1);

    stringReference1.Append(formula1);
    stringReference1.Append(stringCache1);

    seriesText1.Append(stringReference1);

    C.ChartShapeProperties chartShapeProperties1 = new C.ChartShapeProperties();

    A.SolidFill solidFill10 = new A.SolidFill();
    A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "2751A5" };

    solidFill10.Append(rgbColorModelHex1);

    A.Outline outline1 = new A.Outline();
    A.NoFill noFill1 = new A.NoFill();

    outline1.Append(noFill1);
    A.EffectList effectList1 = new A.EffectList();

    chartShapeProperties1.Append(solidFill10);
    chartShapeProperties1.Append(outline1);
    chartShapeProperties1.Append(effectList1);
    C.InvertIfNegative invertIfNegative1 = new C.InvertIfNegative() { Val = false };

    C.DataLabels dataLabels1 = new C.DataLabels();

    C.ChartShapeProperties chartShapeProperties2 = new C.ChartShapeProperties();
    A.NoFill noFill2 = new A.NoFill();

    A.Outline outline2 = new A.Outline();
    A.NoFill noFill3 = new A.NoFill();

    outline2.Append(noFill3);
    A.EffectList effectList2 = new A.EffectList();

    chartShapeProperties2.Append(noFill2);
    chartShapeProperties2.Append(outline2);
    chartShapeProperties2.Append(effectList2);

    C.TextProperties textProperties1 = new C.TextProperties();

    A.BodyProperties bodyProperties1 = new A.BodyProperties() { Rotation = 0, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 38100, TopInset = 19050, RightInset = 38100, BottomInset = 19050, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
    A.ShapeAutoFit shapeAutoFit1 = new A.ShapeAutoFit();

    bodyProperties1.Append(shapeAutoFit1);
    A.ListStyle listStyle1 = new A.ListStyle();

    A.Paragraph paragraph1 = new A.Paragraph();

    A.ParagraphProperties paragraphProperties1 = new A.ParagraphProperties();

    A.DefaultRunProperties defaultRunProperties11 = new A.DefaultRunProperties() { FontSize = 1200, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0 };

    A.SolidFill solidFill11 = new A.SolidFill();
    A.SchemeColor schemeColor10 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

    solidFill11.Append(schemeColor10);
    A.LatinFont latinFont10 = new A.LatinFont() { Typeface = "+mn-lt" };
    A.EastAsianFont eastAsianFont10 = new A.EastAsianFont() { Typeface = "+mn-ea" };
    A.ComplexScriptFont complexScriptFont10 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

    defaultRunProperties11.Append(solidFill11);
    defaultRunProperties11.Append(latinFont10);
    defaultRunProperties11.Append(eastAsianFont10);
    defaultRunProperties11.Append(complexScriptFont10);

    paragraphProperties1.Append(defaultRunProperties11);
    A.EndParagraphRunProperties endParagraphRunProperties1 = new A.EndParagraphRunProperties() { Language = "en-US" };

    paragraph1.Append(paragraphProperties1);
    paragraph1.Append(endParagraphRunProperties1);

    textProperties1.Append(bodyProperties1);
    textProperties1.Append(listStyle1);
    textProperties1.Append(paragraph1);
    C.ShowLegendKey showLegendKey1 = new C.ShowLegendKey() { Val = false };
    C.ShowValue showValue1 = new C.ShowValue() { Val = true };
    C.ShowCategoryName showCategoryName1 = new C.ShowCategoryName() { Val = false };
    C.ShowSeriesName showSeriesName1 = new C.ShowSeriesName() { Val = false };
    C.ShowPercent showPercent1 = new C.ShowPercent() { Val = false };
    C.ShowBubbleSize showBubbleSize1 = new C.ShowBubbleSize() { Val = false };
    C.ShowLeaderLines showLeaderLines1 = new C.ShowLeaderLines() { Val = false };

    C.DLblsExtensionList dLblsExtensionList1 = new C.DLblsExtensionList();

    C.DLblsExtension dLblsExtension1 = new C.DLblsExtension() { Uri = "{CE6537A1-D6FC-4f65-9D91-7224C49458BB}" };
    dLblsExtension1.AddNamespaceDeclaration("c15", "http://schemas.microsoft.com/office/drawing/2012/chart");
    C15.ShowLeaderLines showLeaderLines2 = new C15.ShowLeaderLines() { Val = true };

    C15.LeaderLines leaderLines1 = new C15.LeaderLines();

    C.ChartShapeProperties chartShapeProperties3 = new C.ChartShapeProperties();

    A.Outline outline3 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

    A.SolidFill solidFill12 = new A.SolidFill();

    A.SchemeColor schemeColor11 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
    A.LuminanceModulation luminanceModulation1 = new A.LuminanceModulation() { Val = 35000 };
    A.LuminanceOffset luminanceOffset1 = new A.LuminanceOffset() { Val = 65000 };

    schemeColor11.Append(luminanceModulation1);
    schemeColor11.Append(luminanceOffset1);

    solidFill12.Append(schemeColor11);
    A.Round round1 = new A.Round();

    outline3.Append(solidFill12);
    outline3.Append(round1);
    A.EffectList effectList3 = new A.EffectList();

    chartShapeProperties3.Append(outline3);
    chartShapeProperties3.Append(effectList3);

    leaderLines1.Append(chartShapeProperties3);

    dLblsExtension1.Append(showLeaderLines2);
    dLblsExtension1.Append(leaderLines1);

    dLblsExtensionList1.Append(dLblsExtension1);

    dataLabels1.Append(chartShapeProperties2);
    dataLabels1.Append(textProperties1);
    dataLabels1.Append(showLegendKey1);
    dataLabels1.Append(showValue1);
    dataLabels1.Append(showCategoryName1);
    dataLabels1.Append(showSeriesName1);
    dataLabels1.Append(showPercent1);
    dataLabels1.Append(showBubbleSize1);
    dataLabels1.Append(showLeaderLines1);
    dataLabels1.Append(dLblsExtensionList1);

    C.CategoryAxisData categoryAxisData1 = new C.CategoryAxisData();

    C.StringReference stringReference2 = new C.StringReference();
    C.Formula formula2 = new C.Formula();
    formula2.Text = "Sheet1!$A$2:$A$31";

    C.StringCache stringCache2 = new C.StringCache();
    C.PointCount pointCount2 = new C.PointCount() { Val = (UInt32Value)30U };

    C.StringPoint stringPoint2 = new C.StringPoint() { Index = (UInt32Value)0U };
    C.NumericValue numericValue2 = new C.NumericValue();
    numericValue2.Text = "Category 1";

    stringPoint2.Append(numericValue2);

    C.StringPoint stringPoint3 = new C.StringPoint() { Index = (UInt32Value)1U };
    C.NumericValue numericValue3 = new C.NumericValue();
    numericValue3.Text = "Category 2";

    stringPoint3.Append(numericValue3);

    C.StringPoint stringPoint4 = new C.StringPoint() { Index = (UInt32Value)2U };
    C.NumericValue numericValue4 = new C.NumericValue();
    numericValue4.Text = "Category 3";

    stringPoint4.Append(numericValue4);

    C.StringPoint stringPoint5 = new C.StringPoint() { Index = (UInt32Value)3U };
    C.NumericValue numericValue5 = new C.NumericValue();
    numericValue5.Text = "Category 4";

    stringPoint5.Append(numericValue5);

    stringCache2.Append(pointCount2);
    stringCache2.Append(stringPoint2);
    stringCache2.Append(stringPoint3);
    stringCache2.Append(stringPoint4);
    stringCache2.Append(stringPoint5);

    stringReference2.Append(formula2);
    stringReference2.Append(stringCache2);

    categoryAxisData1.Append(stringReference2);

    C.Values values1 = new C.Values();

    C.NumberReference numberReference1 = new C.NumberReference();
    C.Formula formula3 = new C.Formula();
    formula3.Text = "Sheet1!$B$2:$B$31";

    C.NumberingCache numberingCache1 = new C.NumberingCache();
    C.FormatCode formatCode1 = new C.FormatCode();
    formatCode1.Text = "0.0";
    C.PointCount pointCount3 = new C.PointCount() { Val = (UInt32Value)30U };

    C.NumericPoint numericPoint1 = new C.NumericPoint() { Index = (UInt32Value)0U };
    C.NumericValue numericValue32 = new C.NumericValue();
    numericValue32.Text = "4.3";

    numericPoint1.Append(numericValue32);

    C.NumericPoint numericPoint2 = new C.NumericPoint() { Index = (UInt32Value)1U };
    C.NumericValue numericValue33 = new C.NumericValue();
    numericValue33.Text = "2.5";

    numericPoint2.Append(numericValue33);

    C.NumericPoint numericPoint3 = new C.NumericPoint() { Index = (UInt32Value)2U };
    C.NumericValue numericValue34 = new C.NumericValue();
    numericValue34.Text = "3.5";

    numericPoint3.Append(numericValue34);

    C.NumericPoint numericPoint4 = new C.NumericPoint() { Index = (UInt32Value)3U };
    C.NumericValue numericValue35 = new C.NumericValue();
    numericValue35.Text = "4.5";

    numericPoint4.Append(numericValue35);

    numberingCache1.Append(formatCode1);
    numberingCache1.Append(pointCount3);
    numberingCache1.Append(numericPoint1);
    numberingCache1.Append(numericPoint2);
    numberingCache1.Append(numericPoint3);
    numberingCache1.Append(numericPoint4);

    numberReference1.Append(formula3);
    numberReference1.Append(numberingCache1);

    values1.Append(numberReference1);

    C.BarSerExtensionList barSerExtensionList1 = new C.BarSerExtensionList();

    C.BarSerExtension barSerExtension1 = new C.BarSerExtension() { Uri = "{C3380CC4-5D6E-409C-BE32-E72D297353CC}" };
    barSerExtension1.AddNamespaceDeclaration("c16", "http://schemas.microsoft.com/office/drawing/2014/chart");

    barSerExtensionList1.Append(barSerExtension1);

    barChartSeries1.Append(index1);
    barChartSeries1.Append(order1);
    barChartSeries1.Append(seriesText1);
    barChartSeries1.Append(chartShapeProperties1);
    barChartSeries1.Append(invertIfNegative1);
    barChartSeries1.Append(dataLabels1);
    barChartSeries1.Append(categoryAxisData1);
    barChartSeries1.Append(values1);
    barChartSeries1.Append(barSerExtensionList1);

    C.BarChartSeries barChartSeries2 = new C.BarChartSeries();
    C.Index index2 = new C.Index() { Val = (UInt32Value)1U };
    C.Order order2 = new C.Order() { Val = (UInt32Value)1U };

    C.SeriesText seriesText2 = new C.SeriesText();

    C.StringReference stringReference3 = new C.StringReference();
    C.Formula formula4 = new C.Formula();
    formula4.Text = "Sheet1!$C$1";

    C.StringCache stringCache3 = new C.StringCache();
    C.PointCount pointCount4 = new C.PointCount() { Val = (UInt32Value)1U };

    C.StringPoint stringPoint32 = new C.StringPoint() { Index = (UInt32Value)0U };
    C.NumericValue numericValue62 = new C.NumericValue();
    numericValue62.Text = "Series 2";

    stringPoint32.Append(numericValue62);

    stringCache3.Append(pointCount4);
    stringCache3.Append(stringPoint32);

    stringReference3.Append(formula4);
    stringReference3.Append(stringCache3);

    seriesText2.Append(stringReference3);

    C.ChartShapeProperties chartShapeProperties4 = new C.ChartShapeProperties();

    A.SolidFill solidFill13 = new A.SolidFill();
    A.SchemeColor schemeColor12 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent2 };

    solidFill13.Append(schemeColor12);

    A.Outline outline4 = new A.Outline();
    A.NoFill noFill4 = new A.NoFill();

    outline4.Append(noFill4);
    A.EffectList effectList4 = new A.EffectList();

    chartShapeProperties4.Append(solidFill13);
    chartShapeProperties4.Append(outline4);
    chartShapeProperties4.Append(effectList4);
    C.InvertIfNegative invertIfNegative2 = new C.InvertIfNegative() { Val = false };

    C.DataLabels dataLabels2 = new C.DataLabels();

    C.ChartShapeProperties chartShapeProperties5 = new C.ChartShapeProperties();
    A.NoFill noFill5 = new A.NoFill();

    A.Outline outline5 = new A.Outline();
    A.NoFill noFill6 = new A.NoFill();

    outline5.Append(noFill6);
    A.EffectList effectList5 = new A.EffectList();

    chartShapeProperties5.Append(noFill5);
    chartShapeProperties5.Append(outline5);
    chartShapeProperties5.Append(effectList5);

    C.TextProperties textProperties2 = new C.TextProperties();

    A.BodyProperties bodyProperties2 = new A.BodyProperties() { Rotation = 0, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 38100, TopInset = 19050, RightInset = 38100, BottomInset = 19050, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
    A.ShapeAutoFit shapeAutoFit2 = new A.ShapeAutoFit();

    bodyProperties2.Append(shapeAutoFit2);
    A.ListStyle listStyle2 = new A.ListStyle();

    A.Paragraph paragraph2 = new A.Paragraph();

    A.ParagraphProperties paragraphProperties2 = new A.ParagraphProperties();

    A.DefaultRunProperties defaultRunProperties12 = new A.DefaultRunProperties() { FontSize = 1200, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0 };

    A.SolidFill solidFill14 = new A.SolidFill();
    A.SchemeColor schemeColor13 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

    solidFill14.Append(schemeColor13);
    A.LatinFont latinFont11 = new A.LatinFont() { Typeface = "+mn-lt" };
    A.EastAsianFont eastAsianFont11 = new A.EastAsianFont() { Typeface = "+mn-ea" };
    A.ComplexScriptFont complexScriptFont11 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

    defaultRunProperties12.Append(solidFill14);
    defaultRunProperties12.Append(latinFont11);
    defaultRunProperties12.Append(eastAsianFont11);
    defaultRunProperties12.Append(complexScriptFont11);

    paragraphProperties2.Append(defaultRunProperties12);
    A.EndParagraphRunProperties endParagraphRunProperties2 = new A.EndParagraphRunProperties() { Language = "en-US" };

    paragraph2.Append(paragraphProperties2);
    paragraph2.Append(endParagraphRunProperties2);

    textProperties2.Append(bodyProperties2);
    textProperties2.Append(listStyle2);
    textProperties2.Append(paragraph2);
    C.ShowLegendKey showLegendKey2 = new C.ShowLegendKey() { Val = false };
    C.ShowValue showValue2 = new C.ShowValue() { Val = true };
    C.ShowCategoryName showCategoryName2 = new C.ShowCategoryName() { Val = false };
    C.ShowSeriesName showSeriesName2 = new C.ShowSeriesName() { Val = false };
    C.ShowPercent showPercent2 = new C.ShowPercent() { Val = false };
    C.ShowBubbleSize showBubbleSize2 = new C.ShowBubbleSize() { Val = false };
    C.ShowLeaderLines showLeaderLines3 = new C.ShowLeaderLines() { Val = false };

    C.DLblsExtensionList dLblsExtensionList2 = new C.DLblsExtensionList();

    C.DLblsExtension dLblsExtension2 = new C.DLblsExtension() { Uri = "{CE6537A1-D6FC-4f65-9D91-7224C49458BB}" };
    dLblsExtension2.AddNamespaceDeclaration("c15", "http://schemas.microsoft.com/office/drawing/2012/chart");
    C15.ShowLeaderLines showLeaderLines4 = new C15.ShowLeaderLines() { Val = true };

    C15.LeaderLines leaderLines2 = new C15.LeaderLines();

    C.ChartShapeProperties chartShapeProperties6 = new C.ChartShapeProperties();

    A.Outline outline6 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

    A.SolidFill solidFill15 = new A.SolidFill();

    A.SchemeColor schemeColor14 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
    A.LuminanceModulation luminanceModulation2 = new A.LuminanceModulation() { Val = 35000 };
    A.LuminanceOffset luminanceOffset2 = new A.LuminanceOffset() { Val = 65000 };

    schemeColor14.Append(luminanceModulation2);
    schemeColor14.Append(luminanceOffset2);

    solidFill15.Append(schemeColor14);
    A.Round round2 = new A.Round();

    outline6.Append(solidFill15);
    outline6.Append(round2);
    A.EffectList effectList6 = new A.EffectList();

    chartShapeProperties6.Append(outline6);
    chartShapeProperties6.Append(effectList6);

    leaderLines2.Append(chartShapeProperties6);

    dLblsExtension2.Append(showLeaderLines4);
    dLblsExtension2.Append(leaderLines2);

    dLblsExtensionList2.Append(dLblsExtension2);

    dataLabels2.Append(chartShapeProperties5);
    dataLabels2.Append(textProperties2);
    dataLabels2.Append(showLegendKey2);
    dataLabels2.Append(showValue2);
    dataLabels2.Append(showCategoryName2);
    dataLabels2.Append(showSeriesName2);
    dataLabels2.Append(showPercent2);
    dataLabels2.Append(showBubbleSize2);
    dataLabels2.Append(showLeaderLines3);
    dataLabels2.Append(dLblsExtensionList2);

    C.CategoryAxisData categoryAxisData2 = new C.CategoryAxisData();

    C.StringReference stringReference4 = new C.StringReference();
    C.Formula formula5 = new C.Formula();
    formula5.Text = "Sheet1!$A$2:$A$31";

    C.StringCache stringCache4 = new C.StringCache();
    C.PointCount pointCount5 = new C.PointCount() { Val = (UInt32Value)30U };

    C.StringPoint stringPoint33 = new C.StringPoint() { Index = (UInt32Value)0U };
    C.NumericValue numericValue63 = new C.NumericValue();
    numericValue63.Text = "Category 1";

    stringPoint33.Append(numericValue63);

    C.StringPoint stringPoint34 = new C.StringPoint() { Index = (UInt32Value)1U };
    C.NumericValue numericValue64 = new C.NumericValue();
    numericValue64.Text = "Category 2";

    stringPoint34.Append(numericValue64);

    C.StringPoint stringPoint35 = new C.StringPoint() { Index = (UInt32Value)2U };
    C.NumericValue numericValue65 = new C.NumericValue();
    numericValue65.Text = "Category 3";

    stringPoint35.Append(numericValue65);

    C.StringPoint stringPoint36 = new C.StringPoint() { Index = (UInt32Value)3U };
    C.NumericValue numericValue66 = new C.NumericValue();
    numericValue66.Text = "Category 4";

    stringPoint36.Append(numericValue66);

    stringCache4.Append(pointCount5);
    stringCache4.Append(stringPoint33);
    stringCache4.Append(stringPoint34);
    stringCache4.Append(stringPoint35);
    stringCache4.Append(stringPoint36);

    stringReference4.Append(formula5);
    stringReference4.Append(stringCache4);

    categoryAxisData2.Append(stringReference4);

    C.Values values2 = new C.Values();

    C.NumberReference numberReference2 = new C.NumberReference();
    C.Formula formula6 = new C.Formula();
    formula6.Text = "Sheet1!$C$2:$C$31";

    C.NumberingCache numberingCache2 = new C.NumberingCache();
    C.FormatCode formatCode2 = new C.FormatCode();
    formatCode2.Text = "0.0";
    C.PointCount pointCount6 = new C.PointCount() { Val = (UInt32Value)30U };

    C.NumericPoint numericPoint31 = new C.NumericPoint() { Index = (UInt32Value)0U };
    C.NumericValue numericValue93 = new C.NumericValue();
    numericValue93.Text = "2.4";

    numericPoint31.Append(numericValue93);

    C.NumericPoint numericPoint32 = new C.NumericPoint() { Index = (UInt32Value)1U };
    C.NumericValue numericValue94 = new C.NumericValue();
    numericValue94.Text = "4.4";

    numericPoint32.Append(numericValue94);

    C.NumericPoint numericPoint33 = new C.NumericPoint() { Index = (UInt32Value)2U };
    C.NumericValue numericValue95 = new C.NumericValue();
    numericValue95.Text = "1.8";

    numericPoint33.Append(numericValue95);

    C.NumericPoint numericPoint34 = new C.NumericPoint() { Index = (UInt32Value)3U };
    C.NumericValue numericValue96 = new C.NumericValue();
    numericValue96.Text = "2.8";

    numericPoint34.Append(numericValue96);

    numberingCache2.Append(formatCode2);
    numberingCache2.Append(pointCount6);
    numberingCache2.Append(numericPoint31);
    numberingCache2.Append(numericPoint32);
    numberingCache2.Append(numericPoint33);
    numberingCache2.Append(numericPoint34);

    numberReference2.Append(formula6);
    numberReference2.Append(numberingCache2);

    values2.Append(numberReference2);

    C.BarSerExtensionList barSerExtensionList2 = new C.BarSerExtensionList();

    C.BarSerExtension barSerExtension2 = new C.BarSerExtension() { Uri = "{C3380CC4-5D6E-409C-BE32-E72D297353CC}" };
    barSerExtension2.AddNamespaceDeclaration("c16", "http://schemas.microsoft.com/office/drawing/2014/chart");

    barSerExtensionList2.Append(barSerExtension2);

    barChartSeries2.Append(index2);
    barChartSeries2.Append(order2);
    barChartSeries2.Append(seriesText2);
    barChartSeries2.Append(chartShapeProperties4);
    barChartSeries2.Append(invertIfNegative2);
    barChartSeries2.Append(dataLabels2);
    barChartSeries2.Append(categoryAxisData2);
    barChartSeries2.Append(values2);
    barChartSeries2.Append(barSerExtensionList2);

    C.BarChartSeries barChartSeries3 = new C.BarChartSeries();
    C.Index index3 = new C.Index() { Val = (UInt32Value)2U };
    C.Order order3 = new C.Order() { Val = (UInt32Value)2U };

    C.SeriesText seriesText3 = new C.SeriesText();

    C.StringReference stringReference5 = new C.StringReference();
    C.Formula formula7 = new C.Formula();
    formula7.Text = "Sheet1!$D$1";

    C.StringCache stringCache5 = new C.StringCache();
    C.PointCount pointCount7 = new C.PointCount() { Val = (UInt32Value)1U };

    C.StringPoint stringPoint63 = new C.StringPoint() { Index = (UInt32Value)0U };
    C.NumericValue numericValue123 = new C.NumericValue();
    numericValue123.Text = "Series 3";

    stringPoint63.Append(numericValue123);

    stringCache5.Append(pointCount7);
    stringCache5.Append(stringPoint63);

    stringReference5.Append(formula7);
    stringReference5.Append(stringCache5);

    seriesText3.Append(stringReference5);

    C.ChartShapeProperties chartShapeProperties7 = new C.ChartShapeProperties();

    A.SolidFill solidFill16 = new A.SolidFill();
    A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex() { Val = "FF7D7D" };

    solidFill16.Append(rgbColorModelHex2);

    A.Outline outline7 = new A.Outline();
    A.NoFill noFill7 = new A.NoFill();

    outline7.Append(noFill7);
    A.EffectList effectList7 = new A.EffectList();

    chartShapeProperties7.Append(solidFill16);
    chartShapeProperties7.Append(outline7);
    chartShapeProperties7.Append(effectList7);
    C.InvertIfNegative invertIfNegative3 = new C.InvertIfNegative() { Val = false };

    C.DataLabels dataLabels3 = new C.DataLabels();

    C.ChartShapeProperties chartShapeProperties8 = new C.ChartShapeProperties();
    A.NoFill noFill8 = new A.NoFill();

    A.Outline outline8 = new A.Outline();
    A.NoFill noFill9 = new A.NoFill();

    outline8.Append(noFill9);
    A.EffectList effectList8 = new A.EffectList();

    chartShapeProperties8.Append(noFill8);
    chartShapeProperties8.Append(outline8);
    chartShapeProperties8.Append(effectList8);

    C.TextProperties textProperties3 = new C.TextProperties();

    A.BodyProperties bodyProperties3 = new A.BodyProperties() { Rotation = 0, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 38100, TopInset = 19050, RightInset = 38100, BottomInset = 19050, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
    A.ShapeAutoFit shapeAutoFit3 = new A.ShapeAutoFit();

    bodyProperties3.Append(shapeAutoFit3);
    A.ListStyle listStyle3 = new A.ListStyle();

    A.Paragraph paragraph3 = new A.Paragraph();

    A.ParagraphProperties paragraphProperties3 = new A.ParagraphProperties();

    A.DefaultRunProperties defaultRunProperties13 = new A.DefaultRunProperties() { FontSize = 1200, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0 };

    A.SolidFill solidFill17 = new A.SolidFill();
    A.SchemeColor schemeColor15 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

    solidFill17.Append(schemeColor15);
    A.LatinFont latinFont12 = new A.LatinFont() { Typeface = "+mn-lt" };
    A.EastAsianFont eastAsianFont12 = new A.EastAsianFont() { Typeface = "+mn-ea" };
    A.ComplexScriptFont complexScriptFont12 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

    defaultRunProperties13.Append(solidFill17);
    defaultRunProperties13.Append(latinFont12);
    defaultRunProperties13.Append(eastAsianFont12);
    defaultRunProperties13.Append(complexScriptFont12);

    paragraphProperties3.Append(defaultRunProperties13);
    A.EndParagraphRunProperties endParagraphRunProperties3 = new A.EndParagraphRunProperties() { Language = "en-US" };

    paragraph3.Append(paragraphProperties3);
    paragraph3.Append(endParagraphRunProperties3);

    textProperties3.Append(bodyProperties3);
    textProperties3.Append(listStyle3);
    textProperties3.Append(paragraph3);
    C.ShowLegendKey showLegendKey3 = new C.ShowLegendKey() { Val = false };
    C.ShowValue showValue3 = new C.ShowValue() { Val = true };
    C.ShowCategoryName showCategoryName3 = new C.ShowCategoryName() { Val = false };
    C.ShowSeriesName showSeriesName3 = new C.ShowSeriesName() { Val = false };
    C.ShowPercent showPercent3 = new C.ShowPercent() { Val = false };
    C.ShowBubbleSize showBubbleSize3 = new C.ShowBubbleSize() { Val = false };
    C.ShowLeaderLines showLeaderLines5 = new C.ShowLeaderLines() { Val = false };

    C.DLblsExtensionList dLblsExtensionList3 = new C.DLblsExtensionList();

    C.DLblsExtension dLblsExtension3 = new C.DLblsExtension() { Uri = "{CE6537A1-D6FC-4f65-9D91-7224C49458BB}" };
    dLblsExtension3.AddNamespaceDeclaration("c15", "http://schemas.microsoft.com/office/drawing/2012/chart");
    C15.ShowLeaderLines showLeaderLines6 = new C15.ShowLeaderLines() { Val = true };

    C15.LeaderLines leaderLines3 = new C15.LeaderLines();

    C.ChartShapeProperties chartShapeProperties9 = new C.ChartShapeProperties();

    A.Outline outline9 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

    A.SolidFill solidFill18 = new A.SolidFill();

    A.SchemeColor schemeColor16 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
    A.LuminanceModulation luminanceModulation3 = new A.LuminanceModulation() { Val = 35000 };
    A.LuminanceOffset luminanceOffset3 = new A.LuminanceOffset() { Val = 65000 };

    schemeColor16.Append(luminanceModulation3);
    schemeColor16.Append(luminanceOffset3);

    solidFill18.Append(schemeColor16);
    A.Round round3 = new A.Round();

    outline9.Append(solidFill18);
    outline9.Append(round3);
    A.EffectList effectList9 = new A.EffectList();

    chartShapeProperties9.Append(outline9);
    chartShapeProperties9.Append(effectList9);

    leaderLines3.Append(chartShapeProperties9);

    dLblsExtension3.Append(showLeaderLines6);
    dLblsExtension3.Append(leaderLines3);

    dLblsExtensionList3.Append(dLblsExtension3);

    dataLabels3.Append(chartShapeProperties8);
    dataLabels3.Append(textProperties3);
    dataLabels3.Append(showLegendKey3);
    dataLabels3.Append(showValue3);
    dataLabels3.Append(showCategoryName3);
    dataLabels3.Append(showSeriesName3);
    dataLabels3.Append(showPercent3);
    dataLabels3.Append(showBubbleSize3);
    dataLabels3.Append(showLeaderLines5);
    dataLabels3.Append(dLblsExtensionList3);

    C.CategoryAxisData categoryAxisData3 = new C.CategoryAxisData();

    C.StringReference stringReference6 = new C.StringReference();
    C.Formula formula8 = new C.Formula();
    formula8.Text = "Sheet1!$A$2:$A$31";

    C.StringCache stringCache6 = new C.StringCache();
    C.PointCount pointCount8 = new C.PointCount() { Val = (UInt32Value)30U };

    C.StringPoint stringPoint64 = new C.StringPoint() { Index = (UInt32Value)0U };
    C.NumericValue numericValue124 = new C.NumericValue();
    numericValue124.Text = "Category 1";

    stringPoint64.Append(numericValue124);

    C.StringPoint stringPoint65 = new C.StringPoint() { Index = (UInt32Value)1U };
    C.NumericValue numericValue125 = new C.NumericValue();
    numericValue125.Text = "Category 2";

    stringPoint65.Append(numericValue125);

    C.StringPoint stringPoint66 = new C.StringPoint() { Index = (UInt32Value)2U };
    C.NumericValue numericValue126 = new C.NumericValue();
    numericValue126.Text = "Category 3";

    stringPoint66.Append(numericValue126);

    C.StringPoint stringPoint67 = new C.StringPoint() { Index = (UInt32Value)3U };
    C.NumericValue numericValue127 = new C.NumericValue();
    numericValue127.Text = "Category 4";

    stringPoint67.Append(numericValue127);

    stringCache6.Append(pointCount8);
    stringCache6.Append(stringPoint64);
    stringCache6.Append(stringPoint65);
    stringCache6.Append(stringPoint66);
    stringCache6.Append(stringPoint67);

    stringReference6.Append(formula8);
    stringReference6.Append(stringCache6);

    categoryAxisData3.Append(stringReference6);

    C.Values values3 = new C.Values();

    C.NumberReference numberReference3 = new C.NumberReference();
    C.Formula formula9 = new C.Formula();
    formula9.Text = "Sheet1!$D$2:$D$31";

    C.NumberingCache numberingCache3 = new C.NumberingCache();
    C.FormatCode formatCode3 = new C.FormatCode();
    formatCode3.Text = "0.0";
    C.PointCount pointCount9 = new C.PointCount() { Val = (UInt32Value)30U };

    C.NumericPoint numericPoint61 = new C.NumericPoint() { Index = (UInt32Value)0U };
    C.NumericValue numericValue154 = new C.NumericValue();
    numericValue154.Text = "2";

    numericPoint61.Append(numericValue154);

    C.NumericPoint numericPoint62 = new C.NumericPoint() { Index = (UInt32Value)1U };
    C.NumericValue numericValue155 = new C.NumericValue();
    numericValue155.Text = "2";

    numericPoint62.Append(numericValue155);

    C.NumericPoint numericPoint63 = new C.NumericPoint() { Index = (UInt32Value)2U };
    C.NumericValue numericValue156 = new C.NumericValue();
    numericValue156.Text = "3";

    numericPoint63.Append(numericValue156);

    C.NumericPoint numericPoint64 = new C.NumericPoint() { Index = (UInt32Value)3U };
    C.NumericValue numericValue157 = new C.NumericValue();
    numericValue157.Text = "5";

    numericPoint64.Append(numericValue157);

    numberingCache3.Append(formatCode3);
    numberingCache3.Append(pointCount9);
    numberingCache3.Append(numericPoint61);
    numberingCache3.Append(numericPoint62);
    numberingCache3.Append(numericPoint63);
    numberingCache3.Append(numericPoint64);

    numberReference3.Append(formula9);
    numberReference3.Append(numberingCache3);

    values3.Append(numberReference3);

    C.BarSerExtensionList barSerExtensionList3 = new C.BarSerExtensionList();

    C.BarSerExtension barSerExtension3 = new C.BarSerExtension() { Uri = "{C3380CC4-5D6E-409C-BE32-E72D297353CC}" };
    barSerExtension3.AddNamespaceDeclaration("c16", "http://schemas.microsoft.com/office/drawing/2014/chart");

    barSerExtensionList3.Append(barSerExtension3);

    barChartSeries3.Append(index3);
    barChartSeries3.Append(order3);
    barChartSeries3.Append(seriesText3);
    barChartSeries3.Append(chartShapeProperties7);
    barChartSeries3.Append(invertIfNegative3);
    barChartSeries3.Append(dataLabels3);
    barChartSeries3.Append(categoryAxisData3);
    barChartSeries3.Append(values3);
    barChartSeries3.Append(barSerExtensionList3);

    C.DataLabels dataLabels4 = new C.DataLabels();
    C.ShowLegendKey showLegendKey4 = new C.ShowLegendKey() { Val = false };
    C.ShowValue showValue4 = new C.ShowValue() { Val = false };
    C.ShowCategoryName showCategoryName4 = new C.ShowCategoryName() { Val = false };
    C.ShowSeriesName showSeriesName4 = new C.ShowSeriesName() { Val = false };
    C.ShowPercent showPercent4 = new C.ShowPercent() { Val = false };
    C.ShowBubbleSize showBubbleSize4 = new C.ShowBubbleSize() { Val = false };

    dataLabels4.Append(showLegendKey4);
    dataLabels4.Append(showValue4);
    dataLabels4.Append(showCategoryName4);
    dataLabels4.Append(showSeriesName4);
    dataLabels4.Append(showPercent4);
    dataLabels4.Append(showBubbleSize4);
    C.GapWidth gapWidth1 = new C.GapWidth() { Val = (UInt16Value)64U };
    C.Overlap overlap1 = new C.Overlap() { Val = 100 };
    C.AxisId axisId1 = new C.AxisId() { Val = (UInt32Value)918330120U };
    C.AxisId axisId2 = new C.AxisId() { Val = (UInt32Value)918324000U };

    barChart1.Append(barDirection1);
    barChart1.Append(barGrouping1);
    barChart1.Append(varyColors1);
    barChart1.Append(barChartSeries1);
    barChart1.Append(barChartSeries2);
    barChart1.Append(barChartSeries3);
    barChart1.Append(dataLabels4);
    barChart1.Append(gapWidth1);
    barChart1.Append(overlap1);
    barChart1.Append(axisId1);
    barChart1.Append(axisId2);

    C.CategoryAxis categoryAxis1 = new C.CategoryAxis();
    C.AxisId axisId3 = new C.AxisId() { Val = (UInt32Value)918330120U };

    C.Scaling scaling1 = new C.Scaling();
    C.Orientation orientation1 = new C.Orientation() { Val = C.OrientationValues.MinMax };

    scaling1.Append(orientation1);
    C.Delete delete1 = new C.Delete() { Val = false };
    C.AxisPosition axisPosition1 = new C.AxisPosition() { Val = C.AxisPositionValues.Left };
    C.NumberingFormat numberingFormat1 = new C.NumberingFormat() { FormatCode = "General", SourceLinked = true };
    C.MajorTickMark majorTickMark1 = new C.MajorTickMark() { Val = C.TickMarkValues.None };
    C.MinorTickMark minorTickMark1 = new C.MinorTickMark() { Val = C.TickMarkValues.None };
    C.TickLabelPosition tickLabelPosition1 = new C.TickLabelPosition() { Val = C.TickLabelPositionValues.NextTo };

    C.ChartShapeProperties chartShapeProperties10 = new C.ChartShapeProperties();
    A.NoFill noFill10 = new A.NoFill();

    A.Outline outline10 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

    A.SolidFill solidFill19 = new A.SolidFill();

    A.SchemeColor schemeColor17 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
    A.LuminanceModulation luminanceModulation4 = new A.LuminanceModulation() { Val = 15000 };
    A.LuminanceOffset luminanceOffset4 = new A.LuminanceOffset() { Val = 85000 };

    schemeColor17.Append(luminanceModulation4);
    schemeColor17.Append(luminanceOffset4);

    solidFill19.Append(schemeColor17);
    A.Round round4 = new A.Round();

    outline10.Append(solidFill19);
    outline10.Append(round4);
    A.EffectList effectList10 = new A.EffectList();

    chartShapeProperties10.Append(noFill10);
    chartShapeProperties10.Append(outline10);
    chartShapeProperties10.Append(effectList10);

    C.TextProperties textProperties4 = new C.TextProperties();
    A.BodyProperties bodyProperties4 = new A.BodyProperties() { Rotation = -60000000, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
    A.ListStyle listStyle4 = new A.ListStyle();

    A.Paragraph paragraph4 = new A.Paragraph();

    A.ParagraphProperties paragraphProperties4 = new A.ParagraphProperties();

    A.DefaultRunProperties defaultRunProperties14 = new A.DefaultRunProperties() { FontSize = 1200, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0 };

    A.SolidFill solidFill20 = new A.SolidFill();
    A.SchemeColor schemeColor18 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

    solidFill20.Append(schemeColor18);
    A.LatinFont latinFont13 = new A.LatinFont() { Typeface = "+mn-lt" };
    A.EastAsianFont eastAsianFont13 = new A.EastAsianFont() { Typeface = "+mn-ea" };
    A.ComplexScriptFont complexScriptFont13 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

    defaultRunProperties14.Append(solidFill20);
    defaultRunProperties14.Append(latinFont13);
    defaultRunProperties14.Append(eastAsianFont13);
    defaultRunProperties14.Append(complexScriptFont13);

    paragraphProperties4.Append(defaultRunProperties14);
    A.EndParagraphRunProperties endParagraphRunProperties4 = new A.EndParagraphRunProperties() { Language = "en-US" };

    paragraph4.Append(paragraphProperties4);
    paragraph4.Append(endParagraphRunProperties4);

    textProperties4.Append(bodyProperties4);
    textProperties4.Append(listStyle4);
    textProperties4.Append(paragraph4);
    C.CrossingAxis crossingAxis1 = new C.CrossingAxis() { Val = (UInt32Value)918324000U };
    C.Crosses crosses1 = new C.Crosses() { Val = C.CrossesValues.AutoZero };
    C.AutoLabeled autoLabeled1 = new C.AutoLabeled() { Val = true };
    C.LabelAlignment labelAlignment1 = new C.LabelAlignment() { Val = C.LabelAlignmentValues.Center };
    C.LabelOffset labelOffset1 = new C.LabelOffset() { Val = (UInt16Value)100U };
    C.NoMultiLevelLabels noMultiLevelLabels1 = new C.NoMultiLevelLabels() { Val = false };

    categoryAxis1.Append(axisId3);
    categoryAxis1.Append(scaling1);
    categoryAxis1.Append(delete1);
    categoryAxis1.Append(axisPosition1);
    categoryAxis1.Append(numberingFormat1);
    categoryAxis1.Append(majorTickMark1);
    categoryAxis1.Append(minorTickMark1);
    categoryAxis1.Append(tickLabelPosition1);
    categoryAxis1.Append(chartShapeProperties10);
    categoryAxis1.Append(textProperties4);
    categoryAxis1.Append(crossingAxis1);
    categoryAxis1.Append(crosses1);
    categoryAxis1.Append(autoLabeled1);
    categoryAxis1.Append(labelAlignment1);
    categoryAxis1.Append(labelOffset1);
    categoryAxis1.Append(noMultiLevelLabels1);

    C.ValueAxis valueAxis1 = new C.ValueAxis();
    C.AxisId axisId4 = new C.AxisId() { Val = (UInt32Value)918324000U };

    C.Scaling scaling2 = new C.Scaling();
    C.Orientation orientation2 = new C.Orientation() { Val = C.OrientationValues.MinMax };

    scaling2.Append(orientation2);
    C.Delete delete2 = new C.Delete() { Val = false };
    C.AxisPosition axisPosition2 = new C.AxisPosition() { Val = C.AxisPositionValues.Bottom };

    C.MajorGridlines majorGridlines1 = new C.MajorGridlines();

    C.ChartShapeProperties chartShapeProperties11 = new C.ChartShapeProperties();

    A.Outline outline11 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

    A.SolidFill solidFill21 = new A.SolidFill();

    A.SchemeColor schemeColor19 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
    A.LuminanceModulation luminanceModulation5 = new A.LuminanceModulation() { Val = 15000 };
    A.LuminanceOffset luminanceOffset5 = new A.LuminanceOffset() { Val = 85000 };

    schemeColor19.Append(luminanceModulation5);
    schemeColor19.Append(luminanceOffset5);

    solidFill21.Append(schemeColor19);
    A.Round round5 = new A.Round();

    outline11.Append(solidFill21);
    outline11.Append(round5);
    A.EffectList effectList11 = new A.EffectList();

    chartShapeProperties11.Append(outline11);
    chartShapeProperties11.Append(effectList11);

    majorGridlines1.Append(chartShapeProperties11);
    C.NumberingFormat numberingFormat2 = new C.NumberingFormat() { FormatCode = "0%", SourceLinked = true };
    C.MajorTickMark majorTickMark2 = new C.MajorTickMark() { Val = C.TickMarkValues.None };
    C.MinorTickMark minorTickMark2 = new C.MinorTickMark() { Val = C.TickMarkValues.None };
    C.TickLabelPosition tickLabelPosition2 = new C.TickLabelPosition() { Val = C.TickLabelPositionValues.NextTo };

    C.ChartShapeProperties chartShapeProperties12 = new C.ChartShapeProperties();
    A.NoFill noFill11 = new A.NoFill();

    A.Outline outline12 = new A.Outline();
    A.NoFill noFill12 = new A.NoFill();

    outline12.Append(noFill12);
    A.EffectList effectList12 = new A.EffectList();

    chartShapeProperties12.Append(noFill11);
    chartShapeProperties12.Append(outline12);
    chartShapeProperties12.Append(effectList12);

    C.TextProperties textProperties5 = new C.TextProperties();
    A.BodyProperties bodyProperties5 = new A.BodyProperties() { Rotation = -60000000, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
    A.ListStyle listStyle5 = new A.ListStyle();

    A.Paragraph paragraph5 = new A.Paragraph();

    A.ParagraphProperties paragraphProperties5 = new A.ParagraphProperties();

    A.DefaultRunProperties defaultRunProperties15 = new A.DefaultRunProperties() { FontSize = 1200, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0 };

    A.SolidFill solidFill22 = new A.SolidFill();
    A.SchemeColor schemeColor20 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

    solidFill22.Append(schemeColor20);
    A.LatinFont latinFont14 = new A.LatinFont() { Typeface = "+mn-lt" };
    A.EastAsianFont eastAsianFont14 = new A.EastAsianFont() { Typeface = "+mn-ea" };
    A.ComplexScriptFont complexScriptFont14 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

    defaultRunProperties15.Append(solidFill22);
    defaultRunProperties15.Append(latinFont14);
    defaultRunProperties15.Append(eastAsianFont14);
    defaultRunProperties15.Append(complexScriptFont14);

    paragraphProperties5.Append(defaultRunProperties15);
    A.EndParagraphRunProperties endParagraphRunProperties5 = new A.EndParagraphRunProperties() { Language = "en-US" };

    paragraph5.Append(paragraphProperties5);
    paragraph5.Append(endParagraphRunProperties5);

    textProperties5.Append(bodyProperties5);
    textProperties5.Append(listStyle5);
    textProperties5.Append(paragraph5);
    C.CrossingAxis crossingAxis2 = new C.CrossingAxis() { Val = (UInt32Value)918330120U };
    C.Crosses crosses2 = new C.Crosses() { Val = C.CrossesValues.AutoZero };
    C.CrossBetween crossBetween1 = new C.CrossBetween() { Val = C.CrossBetweenValues.Between };
    C.MajorUnit majorUnit1 = new C.MajorUnit() { Val = 0.2D };

    valueAxis1.Append(axisId4);
    valueAxis1.Append(scaling2);
    valueAxis1.Append(delete2);
    valueAxis1.Append(axisPosition2);
    valueAxis1.Append(majorGridlines1);
    valueAxis1.Append(numberingFormat2);
    valueAxis1.Append(majorTickMark2);
    valueAxis1.Append(minorTickMark2);
    valueAxis1.Append(tickLabelPosition2);
    valueAxis1.Append(chartShapeProperties12);
    valueAxis1.Append(textProperties5);
    valueAxis1.Append(crossingAxis2);
    valueAxis1.Append(crosses2);
    valueAxis1.Append(crossBetween1);
    valueAxis1.Append(majorUnit1);

    C.ShapeProperties shapeProperties1 = new C.ShapeProperties();
    A.NoFill noFill13 = new A.NoFill();

    A.Outline outline13 = new A.Outline();
    A.NoFill noFill14 = new A.NoFill();

    outline13.Append(noFill14);
    A.EffectList effectList13 = new A.EffectList();

    shapeProperties1.Append(noFill13);
    shapeProperties1.Append(outline13);
    shapeProperties1.Append(effectList13);

    plotArea1.Append(layout1);
    plotArea1.Append(barChart1);
    plotArea1.Append(categoryAxis1);
    plotArea1.Append(valueAxis1);
    plotArea1.Append(shapeProperties1);
    C.PlotVisibleOnly plotVisibleOnly1 = new C.PlotVisibleOnly() { Val = true };
    C.DisplayBlanksAs displayBlanksAs1 = new C.DisplayBlanksAs() { Val = C.DisplayBlanksAsValues.Gap };

    C.ExtensionList extensionList1 = new C.ExtensionList();

    C.Extension extension1 = new C.Extension() { Uri = "{56B9EC1D-385E-4148-901F-78D8002777C0}" };
    extension1.AddNamespaceDeclaration("c16r3", "http://schemas.microsoft.com/office/drawing/2017/03/chart");

    extensionList1.Append(extension1);
    C.ShowDataLabelsOverMaximum showDataLabelsOverMaximum1 = new C.ShowDataLabelsOverMaximum() { Val = false };

    chart1.Append(autoTitleDeleted1);
    chart1.Append(plotArea1);
    chart1.Append(plotVisibleOnly1);
    chart1.Append(displayBlanksAs1);
    chart1.Append(extensionList1);
    chart1.Append(showDataLabelsOverMaximum1);

    C.ShapeProperties shapeProperties2 = new C.ShapeProperties();
    A.NoFill noFill15 = new A.NoFill();

    A.Outline outline14 = new A.Outline();
    A.NoFill noFill16 = new A.NoFill();

    outline14.Append(noFill16);
    A.EffectList effectList14 = new A.EffectList();

    shapeProperties2.Append(noFill15);
    shapeProperties2.Append(outline14);
    shapeProperties2.Append(effectList14);

    C.TextProperties textProperties6 = new C.TextProperties();
    A.BodyProperties bodyProperties6 = new A.BodyProperties();
    A.ListStyle listStyle6 = new A.ListStyle();

    A.Paragraph paragraph6 = new A.Paragraph();

    A.ParagraphProperties paragraphProperties6 = new A.ParagraphProperties();

    A.DefaultRunProperties defaultRunProperties16 = new A.DefaultRunProperties() { FontSize = 1200 };

    A.SolidFill solidFill23 = new A.SolidFill();
    A.SchemeColor schemeColor21 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

    solidFill23.Append(schemeColor21);
    A.LatinFont latinFont15 = new A.LatinFont() { Typeface = "Montserrat", Panose = "00000500000000000000", PitchFamily = 2, CharacterSet = 0 };

    defaultRunProperties16.Append(solidFill23);
    defaultRunProperties16.Append(latinFont15);

    paragraphProperties6.Append(defaultRunProperties16);
    A.EndParagraphRunProperties endParagraphRunProperties6 = new A.EndParagraphRunProperties() { Language = "en-US" };

    paragraph6.Append(paragraphProperties6);
    paragraph6.Append(endParagraphRunProperties6);

    textProperties6.Append(bodyProperties6);
    textProperties6.Append(listStyle6);
    textProperties6.Append(paragraph6);

    C.ExternalData externalData1 = new C.ExternalData() { Id = "rId3" };
    C.AutoUpdate autoUpdate1 = new C.AutoUpdate() { Val = false };

    externalData1.Append(autoUpdate1);

    chartSpace1.Append(date19041);
    chartSpace1.Append(editingLanguage1);
    chartSpace1.Append(roundedCorners1);
    chartSpace1.Append(alternateContent1);
    chartSpace1.Append(chart1);
    chartSpace1.Append(shapeProperties2);
    chartSpace1.Append(textProperties6);
    chartSpace1.Append(externalData1);

    chartPart1.ChartSpace = chartSpace1;
}

//using System;
//using System.Linq;
//using System.Xml.Linq;
//using DocumentFormat.OpenXml;
//using DocumentFormat.OpenXml.Drawing.Charts;
//using DocumentFormat.OpenXml.Packaging;

//X();
//static void X()
//{
//    string pptxFilePath = "C:\\Testing\\Template_Creation\\new_test_file\\NewFolder\\Ppt.pptx";

//    using (PresentationDocument presentationDocument = PresentationDocument.Open(pptxFilePath, true))
//    {
//        // Access the presentation part and then the slide part
//        PresentationPart presentationPart = presentationDocument.PresentationPart;
//        SlidePart slidePart = presentationPart.SlideParts.FirstOrDefault();

//        if (slidePart != null)
//        {
//            // Access the first chart part in the slide
//            ChartPart chartPart = slidePart.ChartParts.FirstOrDefault();

//            if (chartPart != null)
//            {
//                // Access the chart object
//                Chart chart = chartPart.ChartSpace.Elements<Chart>().FirstOrDefault();

//                if (chart != null)
//                {
//                    // Assuming it's a bar chart
//                    BarChart barChart = chart.Descendants<BarChart>().FirstOrDefault();

//                    if (barChart != null)
//                    {
//                        string chartXml = barChart.OuterXml;
//                        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
//                        XDocument doc = XDocument.Parse(chartXml);

//                        List<string> categories = doc.Descendants(c + "cat").Elements(c + "strRef").Elements(c + "pt").Select(pt => pt.Element(c + "v").Value).ToList();

//                        // Extract series data (Y-axis values)
//                        List<List<double>> seriesData = doc.Descendants(c + "ser").Select(ser =>
//                            ser.Element(c + "val").Element(c + "numRef").Element(c + "numCache").Elements(c + "pt").Select(pt =>
//                                Convert.ToDouble(pt.Element(c + "v").Value)).ToList()).ToList();

//                        // Display extracted data
//                        Console.WriteLine("Categories:");
//                        foreach (var category in categories)
//                        {
//                            Console.WriteLine(category);
//                        }

//                        Console.WriteLine("\nSeries Data:");
//                        for (int i = 0; i < seriesData.Count; i++)
//                        {
//                            Console.WriteLine("Series " + (i + 1) + ":");
//                            foreach (var value in seriesData[i])
//                            {
//                                Console.WriteLine(value);
//                            }
//                        }

//                            // Check if it's a stacked chart
//                            if (barChart.BarGrouping.Val == BarGroupingValues.PercentStacked)
//                        {
//                            // Modify the chart here
//                            Console.WriteLine("Found 100% stacked bar chart. You can modify it here.");
//                        }
//                        else
//                        {
//                            Console.WriteLine("The chart is not a 100% stacked bar chart.");
//                        }
//                    }
//                    else
//                    {
//                        Console.WriteLine("The chart is not a bar chart.");
//                    }
//                }
//                else
//                {
//                    Console.WriteLine("No chart found in the slide.");
//                }
//            }
//            else
//            {
//                Console.WriteLine("No chart part found in the slide.");
//            }
//        }
//        else
//        {
//            Console.WriteLine("No slide found in the presentation.");
//        }
//    }
//}
