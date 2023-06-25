VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Sub Worksheet_Change(ByVal Target As Range)


Application.EnableEvents = True

    If Not Intersect(Target, Range("A1")) Is Nothing Then
        
        category_api = Cells(1, 1).Value
        
        If category_api = "despesas" Then
            
            Cells(2, 1) = "anoDotacao"
            Cells(3, 1) = "mesDotacao"
            Cells(4, 1) = "codEmpresa"
            Cells(5, 1) = "codOrgao"
            Cells(6, 1) = "codUnidade"
            Cells(7, 1) = "codFuncao"
            Cells(8, 1) = "codSubFuncao"
            Cells(9, 1) = "codProjetoAtividade"
            Cells(10, 1) = "codPrograma"
            Cells(11, 1) = "codCategoria"
            Cells(12, 1) = "codGrupo"
            Cells(13, 1) = "codModalidade"
            Cells(14, 1) = "codElemento"
            Cells(15, 1) = "codFonteRecurso"
            
     
        ElseIf category_api = "unidades" Then
            
            Cells(2, 1) = "codUnidade"
            Cells(3, 1) = "codOrgao"
            Cells(4, 1) = "anoExercicio"
            Range("A5:A15") = ""

       
        ElseIf category_api = "orgaos" Then
     
            Cells(2, 1) = "codOrgao"
            Cells(3, 1) = "anoExercicio"
            Cells(4, 1) = "numPagina"
            Cells(5, 1) = "codEmpresa"
            Range("A6:A15") = ""
       
       
        ElseIf category_api = "liquidacoes" Then
        
            Cells(2, 1) = "codEmpenho"
            Cells(3, 1) = "anoEmpenho"
            Cells(4, 1) = "codEmpresa"
            Range("A5:A15") = ""

      
        End If
          
     End If
     

End Sub



