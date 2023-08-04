Option Explicit

' �Է°��� �޾Ƽ� stdout���� ����ϴ� �Լ�
Function GetInput(prompt)
    Dim input
    input = InputBox(prompt)
    GetInput = input
End Function

' ���� ����
Dim original_name, new_name, file_system, folder, folder_path, folder_name, files, file, last_check, file_list
Dim hasFilesToRename
hasFilesToRename = False

Set file_system = CreateObject("Scripting.fileSystemObject")

' ���� ��ũ��Ʈ�� ��ġ�� ������ ����
folder_path = file_system.GetParentFolderName(WScript.ScriptFullName)
folder_name = file_system.GetBaseName(folder_path)

' ù��°�� �Է¹��� �̸�
original_name = Inputbox("�ٲ� ������ ���۸��� �Է��ϼ���:", "[ "&folder_name&" ]"&"���� �� ���� ���۸� �����")

If original_name = "" Then
    WScript.Echo "�Է°��� ��ȿ���� �ʽ��ϴ�. ��ũ��Ʈ�� �����մϴ�."
    WScript.Quit 1
End If

' �ι�°�� �Է¹��� �̸�
new_name = Inputbox("�ٲ� ������ ���۸��� �Է��ϼ���:","���� �� ���� ���۸� �����")

If original_name = "" Or new_name = "" Then
    WScript.Echo "�Է°��� ��ȿ���� �ʽ��ϴ�. ��ũ��Ʈ�� �����մϴ�."
    WScript.Quit 1
End If

If original_name = new_name Then
    WScript.Echo "�� �Է°��� �����Ͽ� ��ũ��Ʈ�� �����մϴ�."
    WScript.Quit 1
End If

Set folder = file_system.GetFolder(".")

' ������ ���� ��� ���
file_list = "������ ���� ���:" & vbCrLf
For Each file In folder.Files
    If Left(file.Name, Len(original_name)) = original_name Then
        file_list = file_list & file.Name & vbCrLf
        hasFilesToRename = True
    End If
Next

If Not hasFilesToRename Then
    WScript.Echo "[ "&original_name&" ]"&"(��)�� �����ϴ� ������ �����ϴ�. ��ũ��Ʈ�� �����մϴ�."
    WScript.Quit 1
End If

' ��¥�� ���� �̸��� �ٲ� ������ Ȯ��
last_check = MsgBox(file_list & vbCrLf & "��¥�� """ & folder_name & """ ���� ���� �̸��� �����Ͻðڽ��ϱ�?", vbYesNo + vbExclamation, "Ȯ��")
If last_check = vbYes Then
    For Each file In folder.Files
        If Left(file.Name, Len(original_name)) = original_name Then
            file.Name = new_name & Mid(file.Name, Len(original_name) + 1)
        End If
    Next
    WScript.Echo "���� �̸� ������ �Ϸ�Ǿ����ϴ�."
Else
    WScript.Echo "���� �̸� ������ ��ҵǾ����ϴ�."
End If
