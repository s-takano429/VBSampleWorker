Imports System.Reflection

Module ToolsProperty

    Public Function getToolsName() As String
        'Form名設定
        ' 現在実行しているアセンブリ(.exeのアセンブリ)を取得する
        Dim assm As [Assembly] = [Assembly].GetExecutingAssembly()

        ' AssemblyNameを取得する
        Dim name As AssemblyName = assm.GetName()
        Dim version As System.Version = name.Version

        ' 名前とバージョンを取得して表示する
        'Console.WriteLine("{0} {1}", name.Name, name.Version)
        Return name.Name & "  ver." & version.Major & "." & version.Minor & "." & version.Build
    End Function


End Module
