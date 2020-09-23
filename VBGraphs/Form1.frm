VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   368
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   405
   StartUpPosition =   3  'Windows Default
   Begin MSScriptControlCtl.ScriptControl Script1 
      Left            =   3240
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      AllowUI         =   -1  'True
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   3945
      Left            =   120
      ScaleHeight     =   97.026
      ScaleMode       =   0  'User
      ScaleWidth      =   217.927
      TabIndex        =   2
      Top             =   720
      Width           =   5865
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Plot &Graph"
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   4920
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Text            =   "(2*(x^2))-(3*x)-(7)"
      Top             =   270
      Width           =   5175
   End
   Begin VB.Label Label1 
      Caption         =   "Y="
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   300
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Project:           Plotting graphs using simple VB code.
'Programmer:        Hamman W. Samuel WEB: http://samuelsonline.f2g.net/
'Author's Note:     This code is meant to learn how graphs can be plotted using VB. Use it freely!
'Credits:           Thank you Lefteris Eleftherioades http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=46695&lngWId=1
'Requirements:      Command button, text box, picture box, script control (get this one from the Microsoft site)
'Logic:             To draw connected lines, begin a subsequent line at the end point of the previous line
'Other credits:     Partially based and inspired on Relation Grapher 2.0 by David Kyle of N7Soft WEB: http://n7soft.sheddtech.com EMAIL: n7soft@ sheddtech.com
'Legal:             No reverse engineering, piracy or hacking was involved in the compilation of this code!
'Other issues:      If errors occur during loading, re-reference the msscript.ocx (Script control). Its in this project folder

Option Explicit
Dim XMax, YMax, YMin, XMin, XScale, YScale

Sub DrawGraph(GraphPaper As PictureBox, GraphEquation As TextBox, ScriptCtrl As ScriptControl)
Dim X, Y, X1, Y1, X2, Y2, TempX, TempY

Const SpeedMark = 20 'Making this higher increases plotting quality but reduces plotting speed

'Use these to change the size and scale of the graphs
XScale = 10 'Scale in pixels per unit
YScale = 1
XMin = -200 'Minimum values
YMin = -50
XMax = 200 'Maximum value
YMax = 100

'These next lines fit the graph to the graph area and lets you
'manually change the picture box size without affecting the plotting
'Try making the picture box larger or smaller on the form
With GraphPaper
    .ScaleHeight = YMax
    .ScaleWidth = XMax
    YMax = YMax + YMin / 2
    XMax = XMax + XMin / 2
End With

GraphPaper.Cls 'Clear all previous graphs

'This part draws the two axes
GraphPaper.Line (0, ConvertToYUnits(0))-(XMax - XMin, ConvertToYUnits(0))  'X-axis
GraphPaper.Line (XMax, YMax - YMin)-(XMax, 0)    'Y-axis

'This part is for plotting the graph
For X = -SpeedMark To SpeedMark Step 1 / SpeedMark
    
    With ScriptCtrl 'This part evaluates the y-values of the equation that will be used in plotting
        .ExecuteStatement ("X = " & X)
        Y = .Eval(GraphEquation.Text)
    End With
        
    'This is the main part which draws the entire graphs step-by-step by joining tiny lines together, based on the
    'values from the equation
    On Error Resume Next
    If X <> -SpeedMark Then 'Error handling
        X1 = ConvertToXUnits(X)
        X2 = ConvertToXUnits(TempX)
        Y1 = ConvertToYUnits(Y)
        Y2 = ConvertToYUnits(TempY)
        GraphPaper.Line (X1, Y1)-(X2, Y2), RGB(255, 0, 0) 'Make plotted line red
    End If
    On Error GoTo 0
    'These two TempX, TempY are used to set the next point for the line to continue
    'i.e. Line goes from 0,0 to 1,1. Then it goes from 1,1 to 2,2 etc
    TempX = X
    TempY = Y
Next
End Sub

'These two functions convert equation values to values that can be plotted correctly on the picture control
Function ConvertToXUnits(ValueToChange)
ConvertToXUnits = XMax + (ValueToChange * XScale)
End Function

Function ConvertToYUnits(ValueToChange)
ConvertToYUnits = YMax - (ValueToChange * YScale)
End Function

Private Sub Command1_Click()
DrawGraph Picture1, Text1, Script1
End Sub

