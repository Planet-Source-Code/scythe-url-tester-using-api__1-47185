VERSION 5.00
Begin VB.Form FrmUrlTest 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "URL Tester"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   4395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox Text3 
      Height          =   3255
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "FrmUrlTest.frx":0000
      Top             =   1500
      Width           =   2115
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Test"
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Text            =   "www.planet-source-code.com"
      Top             =   360
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1020
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Status"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   780
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Url"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   195
   End
End
Attribute VB_Name = "FrmUrlTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Simple URL Tester using only Api´s
'2003 Scythe

Option Explicit

Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function HttpOpenRequest Lib "wininet.dll" Alias "HttpOpenRequestA" (ByVal hHttpSession As Long, ByVal sVerb As String, ByVal sObjectName As String, ByVal sVersion As String, ByVal sReferer As String, ByVal something As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function HttpQueryInfo Lib "wininet.dll" Alias "HttpQueryInfoA" (ByVal hHttpRequest As Long, ByVal lInfoLevel As Long, ByRef sBuffer As Any, ByRef lBufferLength As Long, ByRef lIndex As Long) As Integer
Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
Private Declare Function HttpSendRequest Lib "wininet.dll" Alias "HttpSendRequestA" (ByVal hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal sOptional As String, ByVal lOptionalLength As Long) As Integer
Private Const HTTP_QUERY_STATUS_CODE = 19
Private Const INTERNET_SERVICE_HTTP = 3
Private Const scUserAgent = "http sample"
Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Private Const INTERNET_FLAG_RELOAD = &H80000000
Private Const INTERNET_FLAG_KEEP_CONNECTION = &H400000





Private Function CheckUrl(Url As String) As Long
 Dim sBuffer         As String * 1024
 Dim lBufferLength   As Long
 Dim hInternetSession As Long
 Dim hInternetConnect As Long
 Dim hHttpOpenRequest As Long
 
 lBufferLength = 1024
 
 'Remove Http if needed
 If UCase(Left$(Url, 7)) = "HTTP://" Then
  Url = Right$(Url, Len(Url) - 7)
 End If
 
 'Open the Internetconnection
 hInternetSession = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)

 If CBool(hInternetSession) = False Then
  CheckUrl = 0
  Exit Function
 End If

 'Connect and get the Status
 hInternetConnect = InternetConnect(hInternetSession, Url, 80, "", "", INTERNET_SERVICE_HTTP, 0, 0)
 hHttpOpenRequest = HttpOpenRequest(hInternetConnect, "GET", "", "HTTP/1.0", vbNullString, 0, INTERNET_FLAG_RELOAD Or INTERNET_FLAG_KEEP_CONNECTION, 0)
 HttpSendRequest hHttpOpenRequest, vbNullString, 0, vbNullString, 0
 HttpQueryInfo hHttpOpenRequest, HTTP_QUERY_STATUS_CODE, ByVal sBuffer, lBufferLength, 0

 CheckUrl = Val(Left$(sBuffer, lBufferLength))
 ' 0 No Connect / Error
 ' 200 OK
 ' 201 Created
 ' 202 Accepted
 ' 204 No Content
 ' 301 Moved Permanently
 ' 302 Moved Temporarily
 ' 304 Not Modified
 ' 400 Bad Request
 ' 401 Unauthorized
 ' 403 Forbidden
 ' 404 Not Found
 ' 500 Internal Server Error
 ' 501 Not Implemented
 ' 502 Bad Gateway
 ' 503 Service Unavailable

 'Close connections
 InternetCloseHandle (hHttpOpenRequest)
 InternetCloseHandle (hInternetSession)
 InternetCloseHandle (hInternetConnect)
End Function

Private Sub Command1_Click()
 Text1 = CheckUrl(Text2)
End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then
  Text1 = CheckUrl(Text2)
 End If
End Sub
