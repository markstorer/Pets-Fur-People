Imports System.Web.HttpContext
Imports UglyMoneyAuction.GlobalFunctions
Imports UglyMoneyAuction.WSWSEmailFunctions
Imports ADODB
Imports System
Imports System.IO
Imports System.Net
Imports System.Text

Public Class FeaturedPet
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim RC As String = ""
        Dim SOAPText As String
        Dim CRLF As String
        CRLF = vbCrLf

        Dim QueryField As Long
        Dim HTTPError As String
        Dim WebResponse As String

        Dim request As WebRequest
        Try
            request = WebRequest.Create("http://www.petango.com/webservices/adoptablesearch/wsAdoptableAnimals.aspx?species=All&sex=A&agegroup=All&onhold=A&orderby=ID&colnum=4&AuthKey=jsintr75cuapk0u4kpr5abjclvs6vcpsjxsvnre6r8kloxhsq8&css=http://sms.petpoint.com/WebServices/adoptablesearch/css/styles.css")
            request.Method = "GET"

            ' Build Soap Call
            'WebResponse = ExecuteSNWebService(request, stURL, SOAPText, "http://www.service-now.com/" & stTable & "/getKeys")

            Dim response As WebResponse
            Dim responseFromServer As String = ""
            request.ContentType = "text/xml;charset=UTF-8"

            Try
                response = request.GetResponse()
                Dim dataStream As IO.Stream = response.GetResponseStream()
                Dim reader As New StreamReader(dataStream)
                WebResponse = reader.ReadToEnd()
                reader.Close()
                dataStream.Close()
                response.Close()
            Catch ex As Exception
                WebResponse = "Failed: " & ex.ToString
                'RC = WSWSInsertLog("InsertUpdate", "Select User Rate - Soap=", SOAPText)
                'RC = WSWSInsertLog("InsertUpdate", "Select User Rate - Result=", ExecuteSNWebService)
            End Try

            If InStr(WebResponse, "tblSearchResults") > 0 Then
                'GetSysIDsByFieldNames = Mid(WebResponse, InStr(WebResponse, "<sys_id>") + 8, InStr(WebResponse, "</sys_id>") - InStr(WebResponse, "<sys_id>") - 8)
                RC = RC
                Dim BasePosition As Integer
                Dim TBodyend As Integer
                Dim TDPosition As Integer
                Dim TDEndPosition As Integer
                Dim FeaturedPetContent As String
                Dim RowCount As Integer
                Dim FinalColumnCount As Integer
                Dim CurrentRowPosition As Integer
                Dim LastRowPosition As Integer
                Dim CurrentColumnPosition As Integer
                Dim FeaturedPetNumber As Long
                Dim FeaturedPetRow As Long
                Dim FeaturedPetColumn As Long

                BasePosition = InStr(WebResponse, "tblSearchResults")
                TBodyend = InStr(BasePosition, WebResponse, "</table")
                RowCount = 0
                CurrentRowPosition = InStr(BasePosition, WebResponse, "<tr")
                While InStr(CurrentRowPosition, WebResponse, "<tr") > 0
                    RowCount = RowCount + 1
                    CurrentRowPosition = InStr(CurrentRowPosition + 1, WebResponse, "<tr")
                    If CurrentRowPosition = 0 Then
                        CurrentRowPosition = Len(WebResponse) - 1
                    Else
                        LastRowPosition = CurrentRowPosition
                    End If
                End While

                FinalColumnCount = 0
                CurrentColumnPosition = InStr(LastRowPosition, WebResponse, "<td")
                While InStr(CurrentColumnPosition, WebResponse, "<td") > 0
                    If InStr(CurrentColumnPosition, WebResponse, "<div") > 0 And InStr(CurrentColumnPosition, WebResponse, "<div") < TBodyend Then
                        FinalColumnCount = FinalColumnCount + 1
                    End If
                    CurrentColumnPosition = InStr(CurrentColumnPosition + 1, WebResponse, "<td")
                    If CurrentColumnPosition = 0 Then
                        CurrentColumnPosition = Len(WebResponse) - 1
                    End If
                End While

                Randomize()

                FeaturedPetNumber = Int(Rnd() * ((RowCount - 1) * 4 + FinalColumnCount) + 0.5)
                FeaturedPetRow = Int((FeaturedPetNumber + 3) / 4)
                FeaturedPetColumn = FeaturedPetNumber - (FeaturedPetRow - 1) * 4

                RowCount = 0
                CurrentRowPosition = BasePosition
                While RowCount < FeaturedPetRow
                    RowCount = RowCount + 1
                    CurrentRowPosition = InStr(CurrentRowPosition + 1, WebResponse, "<tr")
                End While

                FinalColumnCount = 0
                CurrentColumnPosition = CurrentRowPosition
                While FinalColumnCount < FeaturedPetColumn
                    FinalColumnCount = FinalColumnCount + 1
                    CurrentColumnPosition = InStr(CurrentColumnPosition + 1, WebResponse, "<td")
                End While

                TDPosition = CurrentColumnPosition
                'TDPosition = InStr(CurrentColumnPosition, WebResponse, "<td")
                TDEndPosition = InStr(TDPosition, WebResponse, "</td")
                FeaturedPetContent = Mid(WebResponse, TDPosition + 24, TDEndPosition - TDPosition - 24)

                While InStr(FeaturedPetContent, "<div class=""hidden"">") > 0
                    FeaturedPetContent = Left(FeaturedPetContent, InStr(FeaturedPetContent, "<div class=""hidden"">") - 1) & Right(FeaturedPetContent, 8)

                End While
                FeaturedPetContent = Replace(FeaturedPetContent, "href=""", "href=""http://www.petango.com/webservices/adoptablesearch/")
                FeaturedPetCell.Text = FeaturedPetContent

            Else
                'GetSysIDsByFieldNames = ""
                RC = RC
            End If
        Catch es As Exception
            HTTPError = "Failed: " & es.ToString
            'RC = WSWSInsertLog("Get Sys_ID by FieldNames", "https://" & stURL & "/" & stTable & "?WSDL", HTTPError)
        End Try


    End Sub

End Class