VERSION 5.00
Begin VB.Form frmProgress 
   BorderStyle     =   0  'None
   Caption         =   "PortraitFile"
   ClientHeight    =   990
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9720
   Icon            =   "frmProgress.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   990
   ScaleWidth      =   9720
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   840
      Left            =   75
      TabIndex        =   0
      Top             =   0
      Width           =   9570
      Begin VB.Label lblStep 
         Caption         =   "lblStep"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   9255
      End
      Begin VB.Label lblProgress 
         Caption         =   "lblProgress"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   9255
      End
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
