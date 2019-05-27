VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "АРИЗ-У-2010"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15510
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox1_Change()

End Sub

Private Sub ComboBox1_Enter()
With UserForm1.ComboBox1
.AddItem ("2.1 Анализ функции")
.AddItem ("2.2 Функциональный ИКР, ФОП")
.AddItem ("3.1 Противоречие требований")
.AddItem ("3.2 Приёмы и принципы решения противоречий")
.AddItem ("4.1 Элепольная модель системы")
.AddItem ("4.2 Универсальная система стандартов")
.AddItem ("5.1 Ресурсный ИКР. Противоречие свойств")
.AddItem ("5.2 ИКР свойств. Мобилизация ресурсов")
.AddItem ("6.1 Критика идей. Изменение задач")
.AddItem ("6.2 Замена и обобщение задач. Смена аспекта")
End With
End Sub
