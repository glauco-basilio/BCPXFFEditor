VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--- clsControl ---

Option Explicit

Public LeftIndent As Long
Public RightIndent As Long
Public TopIndent As Long
Public BottomIndent As Long
Public Width As Long
Public Height As Long
Public style As enLayoutStyle
Public ctrl As Control

