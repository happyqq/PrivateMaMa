Attribute VB_Name = "modMachine"
Public Function GetDriveSerialNumber(Optional ByVal DriveLetter As String) As Long
      
          Dim fso As Object, Drv As Object
          
          'Create a FileSystemObject object
          Set fso = CreateObject("Scripting.FileSystemObject")
          
          'Assign the current drive letter if not specified
          If DriveLetter <> "" Then
              Set Drv = fso.GetDrive(DriveLetter)
          Else
              Set Drv = fso.GetDrive(fso.GetDriveName(App.Path))
          End If
      
          With Drv
              If .IsReady Then
                  DriveSerial = Abs(.SerialNumber)
              Else    '"Drive Not Ready!"
                  DriveSerial = -1
              End If
          End With
          
          'Clean up
          Set Drv = Nothing
          Set fso = Nothing
          
          GetDriveSerialNumber = DriveSerial
          
End Function
      

