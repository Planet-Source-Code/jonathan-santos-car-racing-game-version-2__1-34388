Attribute VB_Name = "Module1"
'Author: Jonathan Santos
'Please vote for me
'Feel free to edit to edit the code just put my name somewhere
'Questions or comentaries: JJonathann@msn.com




Public Type Car ' Declares a car properties
Accelerarion As Single
Speed As Single
Maxspeed As Single
End Type

Public viper As Car
Public corvette65 As Car
Public go As Boolean
Public FinishLine As Boolean

Public num As Byte


Public Function Start()
FL = False
go = False
num = num + 1
If num = 2 Then
Form1.Label1.Caption = "Go!"
num = 0
go = True
Form1.Timer1.Interval = 0
End If

End Function



Public Function ForwardV()
If viper.Speed >= viper.Maxspeed Then
viper.Speed = viper.Maxspeed
Form1.yellow.Left = Form1.yellow.Left + viper.Speed
Form1.SpeedV.Caption = "Speed " & CStr(viper.Speed) & " MPH"
Else
viper.Speed = viper.Speed + viper.Accelerarion
Form1.yellow.Left = Form1.yellow.Left + viper.Speed
Form1.SpeedV.Caption = "Speed " & CStr(viper.Speed) & " MPH"
End If

'Checks which car gets to the finish line first:
If ((Form1.yellow.Left + Form1.yellow.Width) >= Form1.FL.X1) And (FinishLine = False) Then
FinishLine = True
go = False
Form1.Timer2.Interval = 0
Form1.Timer3.Interval = 0
Form1.Label2.Caption = "Viper Won"
Form1.Label2.Visible = True

End If
End Function

Public Function ForwardC()
If corvette65.Speed >= corvette65.Maxspeed Then
corvette65.Speed = corvette65.Maxspeed
Form1.green.Left = Form1.green.Left + corvette65.Speed
Form1.SpeedC.Caption = "Speed " & CStr(corvette65.Speed) & " MPH"
Else
corvette65.Speed = corvette65.Speed + corvette65.Accelerarion
Form1.green.Left = Form1.green.Left + corvette65.Speed
Form1.SpeedC.Caption = "Speed " & CStr(corvette65.Speed) & " MPH"
End If

'Checks which car gets to the finish line first:
If ((Form1.green.Left + Form1.green.Width) >= Form1.FL.X1) And FinishLine = False Then
FinishLine = True
go = False
Form1.Timer2.Interval = 0
Form1.Timer3.Interval = 0
Form1.Label2.Caption = "GTR - 1 WON!"
Form1.Label2.Visible = True

End If
End Function

Public Function FormSize()
'Makes the form screen wide
Form1.Width = Screen.Width
Form1.Left = 0
End Function

Public Function Set_Car_Properties()
'Change the values if you want
viper.Maxspeed = 500
viper.Accelerarion = 5
corvette65.Maxspeed = 500
corvette65.Accelerarion = 5
End Function


