Option Explicit

' 입력값을 받아서 stdout으로 출력하는 함수
Function GetInput(prompt)
    Dim input
    input = InputBox(prompt)
    GetInput = input
End Function

' 변수 선언
Dim original_name, new_name, file_system, folder, folder_path, folder_name, files, file, last_check, file_list
Dim hasFilesToRename
hasFilesToRename = False

Set file_system = CreateObject("Scripting.fileSystemObject")

' 현재 스크립트가 위치한 폴더명 추출
folder_path = file_system.GetParentFolderName(WScript.ScriptFullName)
folder_name = file_system.GetBaseName(folder_path)

' 첫번째로 입력받을 이름
original_name = Inputbox("바꿀 파일의 시작명을 입력하세요:", "[ "&folder_name&" ]"&"폴더 내 파일 시작명 변경기")

If original_name = "" Then
    WScript.Echo "입력값이 유효하지 않습니다. 스크립트를 종료합니다."
    WScript.Quit 1
End If

' 두번째로 입력받을 이름
new_name = Inputbox("바뀔 파일의 시작명을 입력하세요:","폴더 내 파일 시작명 변경기")

If original_name = "" Or new_name = "" Then
    WScript.Echo "입력값이 유효하지 않습니다. 스크립트를 종료합니다."
    WScript.Quit 1
End If

If original_name = new_name Then
    WScript.Echo "두 입력값이 동일하여 스크립트를 종료합니다."
    WScript.Quit 1
End If

Set folder = file_system.GetFolder(".")

' 변경할 파일 목록 출력
file_list = "변경할 파일 목록:" & vbCrLf
For Each file In folder.Files
    If Left(file.Name, Len(original_name)) = original_name Then
        file_list = file_list & file.Name & vbCrLf
        hasFilesToRename = True
    End If
Next

If Not hasFilesToRename Then
    WScript.Echo "[ "&original_name&" ]"&"(으)로 시작하는 파일이 없습니다. 스크립트를 종료합니다."
    WScript.Quit 1
End If

' 진짜로 파일 이름을 바꿀 것인지 확인
last_check = MsgBox(file_list & vbCrLf & "진짜로 """ & folder_name & """ 내의 파일 이름을 변경하시겠습니까?", vbYesNo + vbExclamation, "확인")
If last_check = vbYes Then
    For Each file In folder.Files
        If Left(file.Name, Len(original_name)) = original_name Then
            file.Name = new_name & Mid(file.Name, Len(original_name) + 1)
        End If
    Next
    WScript.Echo "파일 이름 변경이 완료되었습니다."
Else
    WScript.Echo "파일 이름 변경이 취소되었습니다."
End If
