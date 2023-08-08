# ArgParse
📋 Simple command-line argument parser for Visual Basic 6

### Dependencies

- Microsoft Scripting Runtime (`sccrun.dll`)

### Usage

#### UNIX-style parsing

```vb
Dim Args As New ArgParser
With Args
    .MarkAsOption "option1", "option2", "option3"
    '--option1, --option2 and --option3 will be marked as "options" with values

    .SetAlias "option3", "o3"
    'You can also set shorter aliases for flags and options

    .Parse "verb1 verb2 ""C:\path\to\example folder"" --option1 value1 --option2=value2 -o3 value3 --flag1 --flag2"
    'Parse the command-line arguments.
    'You can pass in "Command$" in pratical use.
    
    'Show results
    Dim PlainArg As Variant
    For Each PlainArg In .PlainArgs
        Debug.Print PlainArg
    Next

    Debug.Print .Options("option3")
    Debug.Print .FlagEnabled("flag1")
End With
```

#### DOS-style parsing

ArgParse also supports MS-DOS style options which uses slashes as marks and case-insensitive.

```vb
Dim Args As New ArgParser
With Args
    .OptionsStyle = DOS
    .SetAlias "Option3", "O3"
    .Parse "Verb1 Verb2 ""C:\path\to\example folder"" /Option1 Value1 /Option2:Value2 /O3 Value3 /Flag1 /Flag2"

    Dim PlainArg As Variant
    For Each PlainArg In .PlainArgs
        Debug.Print PlainArg
    Next

    Debug.Print .Options("option3")
    Debug.Print .FlagEnabled("flag1")
    'This will work because DOS mode is case-insensitive
End With
```

#### Walking through plain arguments

Calling `Args.NextArg` and `Args.ThisArg` will let you walking through all the plain arguments.

This will be useful in `Select Case ...` statement to select the verb.

```vb
Dim Args As New ArgParser
Args.Parse "bisect bad"
Select Case Args.NextArg
    Case "init"
    Case "clone"
    Case "commit"
    Case "bisect"
    Select Case Args.NextArg
        Case "start"
        Case "good"
        Case "bad"
        Case Else
        '...
    End Select
    Case Else
        Debug.Print "Unsupported operation: " + Args.ThisArg
End Select
```

