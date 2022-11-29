' Substitui os valores da coluna 
Function Replace_Column(Column As String, What As String, Replacement As String)
    Columns(Column).Replace What:=What, Replacement:=Replacement
End Function

' Retorna a extensão do arquivo com base no caminho e nome
Public Function Excel_File_Extension(FilePath, Filename)
    Dim extFind As String
    Dim sFile As String

    sFile = Dir(FilePath & Filename & "*")
    extFind = Right$(sFile, Len(sFile) - InStrRev(sFile, "."))
    Excel_File_Extension = FilePath & Filename & "." & extFind
End Function