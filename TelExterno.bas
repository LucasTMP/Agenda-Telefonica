Attribute VB_Name = "TelExterno"
Public Sub TelExt()


Select Case TabStrip3.SelectedItem

       Case Is = TabStrip3.Tabs(1)

        rs.CursorLocation = adUseClient
        rs.Open "Select NOME, NUMERO from Externo WHERE NOME like 'A%'", db, 3, 3
        Set tel.DataSource = rs
        Set rs = Nothing
    
      Case Is = TabStrip3.Tabs(2)
      
        rs.CursorLocation = adUseClient
        rs.Open "Select NOME, NUMERO from Externo WHERE NOME like 'B%'", db, 3, 3
        Set tel.DataSource = rs
        Set rs = Nothing
        
    
    Case Is = TabStrip3.Tabs(3)
        
        rs.CursorLocation = adUseClient
        rs.Open "Select NOME, NUMERO from Externo WHERE NOME like 'C%'", db, 3, 3
        Set tel.DataSource = rs
        Set rs = Nothing
        
    Case Is = TabStrip3.Tabs(4)

        rs.CursorLocation = adUseClient
        rs.Open "Select NOME, NUMERO from Externo WHERE NOME like 'D%'", db, 3, 3
        Set tel.DataSource = rs
        Set rs = Nothing

    Case Is = TabStrip3.Tabs(5)

        rs.CursorLocation = adUseClient
        rs.Open "Select NOME, NUMERO from Externo WHERE NOME like 'E%'", db, 3, 3
        Set tel.DataSource = rs
        Set rs = Nothing

    Case Is = TabStrip3.Tabs(6)
    
        rs.CursorLocation = adUseClient
        rs.Open "Select NOME, NUMERO from Externo WHERE NOME like 'F%'", db, 3, 3
        Set tel.DataSource = rs
        Set rs = Nothing

    Case Is = TabStrip3.Tabs(7)
    
        rs.CursorLocation = adUseClient
        rs.Open "Select NOME, NUMERO from Externo WHERE NOME like 'G%'", db, 3, 3
        Set tel.DataSource = rs
        Set rs = Nothing
           
    Case Is = TabStrip3.Tabs(8)

        rs.CursorLocation = adUseClient
        rs.Open "Select NOME, NUMERO from Externo WHERE NOME like 'H%'", db, 3, 3
        Set tel.DataSource = rs
        Set rs = Nothing
    
    Case Is = TabStrip3.Tabs(9)

        rs.CursorLocation = adUseClient
        rs.Open "Select NOME, NUMERO from Externo WHERE NOME like 'I%'", db, 3, 3
        Set tel.DataSource = rs
        Set rs = Nothing
    
    Case Is = TabStrip3.Tabs(10)

        rs.CursorLocation = adUseClient
        rs.Open "Select NOME, NUMERO from Externo WHERE NOME like 'J%'", db, 3, 3
        Set tel.DataSource = rs
        Set rs = Nothing
    
    Case Is = TabStrip3.Tabs(11)

        rs.CursorLocation = adUseClient
        rs.Open "Select NOME, NUMERO from Externo WHERE NOME like 'K%'", db, 3, 3
        Set tel.DataSource = rs
        Set rs = Nothing
        
    Case Is = TabStrip3.Tabs(12)
    
        rs.CursorLocation = adUseClient
        rs.Open "Select NOME, NUMERO from Externo WHERE NOME like 'L%'", db, 3, 3
        Set tel.DataSource = rs
        Set rs = Nothing
        
    Case TabStrip3.Tabs(13)

        rs.CursorLocation = adUseClient
        rs.Open "Select NOME, NUMERO from Externo WHERE NOME like 'M%'", db, 3, 3
        Set tel.DataSource = rs
        Set rs = Nothing
    
    Case Is = TabStrip3.Tabs(14)

        rs.CursorLocation = adUseClient
        rs.Open "Select NOME, NUMERO from Externo WHERE NOME like 'N%'", db, 3, 3
        Set tel.DataSource = rs
        Set rs = Nothing
    
    Case Is = TabStrip3.Tabs(15)
    
        rs.CursorLocation = adUseClient
        rs.Open "Select NOME, NUMERO from Externo WHERE NOME like 'O%'", db, 3, 3
        Set tel.DataSource = rs
        Set rs = Nothing
    
    Case Is = TabStrip3.Tabs(16)
    
        rs.CursorLocation = adUseClient
        rs.Open "Select NOME, NUMERO from Externo WHERE NOME like 'P%'", db, 3, 3
        Set tel.DataSource = rs
        Set rs = Nothing
        
    Case Is = TabStrip3.Tabs(17)
    
        rs.CursorLocation = adUseClient
        rs.Open "Select NOME, NUMERO from Externo WHERE NOME like 'Q%'", db, 3, 3
        Set tel.DataSource = rs
        Set rs = Nothing
        
    Case Is = TabStrip3.Tabs(18)

        rs.CursorLocation = adUseClient
        rs.Open "Select NOME, NUMERO from Externo WHERE NOME like 'R%'", db, 3, 3
        Set tel.DataSource = rs
        Set rs = Nothing
    
    Case Is = TabStrip3.Tabs(19)

        rs.CursorLocation = adUseClient
        rs.Open "Select NOME, NUMERO from Externo WHERE NOME like 'S%'", db, 3, 3
        Set tel.DataSource = rs
        Set rs = Nothing
    
    Case Is = TabStrip3.Tabs(20)

        rs.CursorLocation = adUseClient
        rs.Open "Select NOME, NUMERO from Externo WHERE NOME like 'T%'", db, 3, 3
        Set tel.DataSource = rs
        Set rs = Nothing
        
    Case Is = TabStrip3.Tabs(21)

        rs.CursorLocation = adUseClient
        rs.Open "Select NOME, NUMERO from Externo WHERE NOME like 'U%'", db, 3, 3
        Set tel.DataSource = rs
        Set rs = Nothing
        
    Case Is = TabStrip3.Tabs(22)
    
        rs.CursorLocation = adUseClient
        rs.Open "Select NOME, NUMERO from Externo WHERE NOME like 'V%'", db, 3, 3
        Set tel.DataSource = rs
        Set rs = Nothing
        
    Case Is = TabStrip3.Tabs(23)
    
        rs.CursorLocation = adUseClient
        rs.Open "Select NOME, NUMERO from Externo WHERE NOME like 'W%'", db, 3, 3
        Set tel.DataSource = rs
        Set rs = Nothing
        
    Case Is = TabStrip3.Tabs(24)
    
        rs.CursorLocation = adUseClient
        rs.Open "Select NOME, NUMERO from Externo WHERE NOME like 'X%'", db, 3, 3
        Set tel.DataSource = rs
        Set rs = Nothing
         
    Case Is = TabStrip3.Tabs(25)

        rs.CursorLocation = adUseClient
        rs.Open "Select NOME, NUMERO from Externo WHERE NOME like 'Y%'", db, 3, 3
        Set tel.DataSource = rs
        Set rs = Nothing
        
    Case Is = TabStrip3.Tabs(26)

        rs.CursorLocation = adUseClient
        rs.Open "Select NOME, NUMERO from Externo WHERE NOME like 'Z%'", db, 3, 3
        Set tel.DataSource = rs
        Set rs = Nothing

    Case Else
    
         MsgBox "Erro Inesperado."
         
         
End Select


End Sub


