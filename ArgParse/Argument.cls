VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Argument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Enums
Public Enum ArgumentTypes
    PlainArg   'Verb / string / other arguments
    OptionArg  '--option=value
    FlagArg    '--flag
End Enum

Public Enum ArgumentValueTypes
    StringValue
    IntegerValue
    FloatValue
End Enum

Public ArgType As ArgumentTypes
Public ValueType As ArgumentValueTypes
Public Key As String      'Optional
Public Value As Variant

