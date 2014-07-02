Imports System.IO
Imports System.Linq
Imports System.Xml.XPath
Imports System.Decimal

Public Class FeedBuilder
    Public Shared Sub ReadPartFeed()
        Dim partFeed As XDocument = XDocument.Load(Settings.PartFeed)
        Dim pageFeed As XDocument = XDocument.Load(Settings.WebSite & "/pages.xml")
        Dim catFeed As XDocument = XDocument.Load(Settings.WebSite & "/cat.xml")

        Dim deliveryPrice As String = GetDelivery(partFeed)

        Dim joinPart As IEnumerable = From j In (
                                      From partPart In partFeed.Descendants("PART")
                                      Join pagePart In pageFeed.Descendants("page")
                                      On partPart.Element("PARTNAME").Value Equals pagePart.Attribute("part")
                                      Join catPart In catFeed.Descendants("cat")
                                      On catPart.@id Equals pagePart.@id
                                      Where partPart.@DELIVERY Is Nothing
                                      Select New GoogleProduct With {
                                        .Id = partPart.<PARTNAME>.Value,
                                        .Title = partPart.<PARTDES>.Value,
                                        .Description = LTrim(cleanHTML.Clean(partPart.<PARTREMARK>.Value)),
                                        .GoogleCategory = partPart.<GOOGLECATEGORY>.Value,
                                        .ProductCategory = partPart.<GOOGLECATEGORY>.Value,
                                        .Link = Settings.WebSite & "/" & pagePart.@id,
                                        .ImageLink = FindImage(GetImgSection(partPart.<PARTNAME>.Value, pageFeed).ToString(), _
                                                               partPart.<PARTNAME>.Value),
                                        .Condition = "new",
                                        .Availability = FindAvailability(partPart.<AVAILABLE>.Value.ToString()),
                                        .SalePrice = If(GetWasprice(partPart.<PARTNAME>.Value, partFeed) <> "",
                                                        FindPrice(partPart.Element("PRICE")),
                                                        ""),
                                        .Price = If(GetWasprice(partPart.<PARTNAME>.Value, partFeed) <> "",
                                                    GetWasprice(partPart.<PARTNAME>.Value, partFeed),
                                                    FindPrice(partPart.Element("PRICE"))),
                                        .Brand = GetBrand(partPart.<PARTNAME>.Value, partFeed),
                                        .Mpn = GetMPN(partPart.<PARTNAME>.Value, partFeed),
                                        .Delivery = deliveryPrice,
                                        .ShowOnMenu = catPart.@showonmenu
                                            })
                                    Where j.ImageLink <> "" And j.ShowOnMenu = True _
                                    And j.Brand <> "" And j.Mpn <> ""

        Dim ns As XNamespace = "http://base.google.com/ns/1.0"

        If joinPart Is Nothing Then
            Exit Sub
        End If

        'The below query should be made more generic if another customer wants this.

        Dim oDoc As New XDocument(
            New XDeclaration("1.0", "", ""),
            New XElement("rss",
                         New XElement("channel",
                                      New XElement("title", "Window Security Direct"),
                                      New XElement("link", "http://www.windowsecuritydirect.co.uk"),
                                      New XElement("description", "Window Security Direct"),
                                      From i In joinPart
                                      Order By i.Id
                                      Select New XElement("item",
                                        New XElement("title", i.Title),
                                        New XElement("description", i.Description),
                                        New XElement("link", i.Link),
                                        New XElement(ns + "id", i.Id),
                                        New XElement(ns + "condition", i.Condition),
                                        New XElement(ns + "price", i.Price),
                                        If(i.SalePrice <> "",
                                           New XElement(ns + "sale_price", i.SalePrice), Nothing),
                                        New XElement(ns + "availability", i.Availability),
                                        New XElement(ns + "image_link", i.ImageLink),
                                        New XElement(ns + "shipping",
                                                     New XElement(ns + "country", "UK"),
                                                     New XElement(ns + "service", "Standard"),
                                                     New XElement(ns + "price", i.Delivery)),
                                        New XElement(ns + "brand", i.Brand),
                                        New XElement(ns + "mpn", i.Mpn),
                                        If(i.GoogleCategory <> "",
                                           New XElement(ns + "google_product_category"), Nothing))),
            New XAttribute(XNamespace.Xmlns + "g", "http://base.google.com/ns/1.0"),
            New XAttribute("version", "2.0"))
                            )

        oDoc.Save("googlefeed.xml")
    End Sub

    Private Shared Function GetDelivery(ByVal partFeed As XDocument)
        Dim pName As String = partFeed.XPathSelectElement("BASKET/CURRENCY/CODE").@DEFDEL
        Dim y = From a In partFeed.Descendants("PART")
                Where a.<PARTNAME>.Value = pName
                Select a.<PRICE>.<CURRENCY>.<FAMILY>.@PRICE

        Dim delPrice As String = Round(CDec(y(0)) * 1.2, 2, MidpointRounding.AwayFromZero)

        Return delPrice & " GBP"

    End Function

    Private Shared Function GetMPN(ByVal partname As String,
                                     ByVal partFeed As XDocument)
        Dim x As String = String.Format("//PART[PARTNAME=""{0}""]/SPECS/SPEC[@DES=""Manufacturer Code""]", partname)
        Try
            Return partFeed.XPathSelectElement(x).@VALUE
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Private Shared Function GetBrand(ByVal partname As String,
                                     ByVal partFeed As XDocument)
        Dim x As String = String.Format("//PART[PARTNAME=""{0}""]/SPECS/SPEC[@DES=""Brand""]", partname)
        Try
            Return partFeed.XPathSelectElement(x).@VALUE
        Catch ex As Exception
            Return ""
        End Try

    End Function

    Private Shared Function FindImage(ByVal text As String, ByVal partname As String)
        Try
            Dim startPos As Integer = text.IndexOf(String.Format("src={0}", Chr(34))) + 5
            Dim endPos As Integer = text.IndexOf(String.Format("{0}", Chr(34)), startPos)
            Dim j = text.Substring(startPos, endPos - startPos)
            Return j
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Private Shared Function FindAvailability(ByVal avail As String)
        If avail = "Low Stock" Or avail = "In stock" Then
            Return "in stock"
        Else
            Return "out of stock"
        End If
    End Function

    Private Shared Function FindPrice(ByVal price As XElement)
        Dim nowprice, vat As String
        Try
            vat = price.Element("CURRENCY").Attribute("TAXRATE")
            nowprice = price.Element("CURRENCY").Element("BREAK").@PRICE

            nowprice = Round((Decimal.Parse(nowprice) * If(Decimal.Parse(vat) = 0.0, 1.0, (1.0 + Decimal.Parse(vat) / 100))), 2, MidpointRounding.AwayFromZero).ToString()
            If nowprice <> "" Then
                Return nowprice & " GBP"
            Else
                Throw New NullReferenceException
            End If
        Catch ex As Exception
            Return ""
        End Try

    End Function

    Public Shared Function GetImgSection(ByVal partname As String,
                               ByVal pageFeed As XDocument)
        Dim x As String = String.Format("//page[@part=""{0}""]/section[@placeholder=""main_product_images""]", partname)
        Return pageFeed.XPathSelectElement(x).@html
    End Function

    Public Shared Function GetWasprice(ByVal partname As String,
                                       ByVal partFeed As XDocument)
        Dim x As String = String.Format("//PART[PARTNAME=""{0}""]/SPECS/SPEC[@DES=""wasprice""]", partname)
        Try
            Dim wasprice As String = partFeed.XPathSelectElement(x).@VALUE
            If wasprice <> "" Then
                Return wasprice & " GBP"
            Else
                Throw New NullReferenceException
            End If
        Catch ex As NullReferenceException
            Return ""
        End Try
    End Function
End Class

Public Class GoogleProduct
    Public Property Id() As String
    Public Property Title() As String
    Public Property Description() As String
    Public Property GoogleCategory() As String
    Public Property ProductCategory() As String
    Public Property Link() As String
    Public Property ImageLink() As String
    Public Property Condition() As String = "new"
    Public Property Availability() As String
    Public Property Price() As String
    Public Property SalePrice() As String
    Public Property Brand() As String
    Public Property Mpn() As String
    Public Property Delivery() As String
    Public Property ShowOnMenu() As Boolean
End Class

Public Class Settings
    Shared x As XDocument = XDocument.Load("settings.xml")

    Public Shared PartFeed As String = x.XPathSelectElement("settings/partfeed").Value
    Public Shared Website As String = x.XPathSelectElement("settings/website").Value

    Shared Sub CreateSettings()
        If Not File.Exists("settings.xml") Then
            File.Create("settings.xml")
        End If
    End Sub
End Class
