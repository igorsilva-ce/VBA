' Retorna a extensão do arquivo com base no caminho e no nome
Public Function File_Extension(FilePath, Filename)
    Dim extFind As String
    Dim sFile As String

    sFile = Dir(FilePath & Filename & "*")
    extFind = Right$(sFile, Len(sFile) - InStrRev(sFile, "."))
    File_Extension = FilePath & Filename & "." & extFind
End Function