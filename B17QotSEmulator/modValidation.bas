'******************************************************************************
' modValidation.bas
'
' @author Preston V. McMurry III, http://www.prestonm.com
' @copyright (C) Copyright 2002, 2010 by Preston V. McMurry III, http://www.prestonm.com
'
' *****************************************************************************
'
' This file is part of B17QotS, the "B-17: Queen of the Skies" Emulator.
'
' B17QotS is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' B17QotS is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with B17QotS. If not, see <http://www.gnu.org/licenses/>.
'******************************************************************************

Attribute VB_Name = "modValidation"
Option Explicit

'******************************************************************************
' ValidateRequiredInput
'
' INPUT:  A control.
'
' OUTPUT: n/a
'
' RETURN: True if the field contains data, otherwise false.
'
' NOTES:  IsBlank() wrapper which requires a field contain data.
'******************************************************************************
Public Function ValidateRequiredInput(ctlField As Control) As Boolean
    ValidateRequiredInput = True
    
    If IsBlank(ctlField) = True Then
        MsgBox "Required field is blank.", vbExclamation
        ctlField.SetFocus
        ValidateRequiredInput = False
    End If
End Function

'******************************************************************************
' IsBlank
'
' INPUT:  A control.
'
' OUTPUT: n/a
'
' RETURN: True if it is blank (empty), otherwise false.
'
' NOTES:  n/a
'******************************************************************************
Public Function IsBlank(ctlField As Control) As Boolean
    Dim intFieldLen As Integer
    Dim intCounter As Integer
    
    IsBlank = True
    
    intFieldLen = Len(ctlField.Text)
    
    If intFieldLen = 0 Then
        Exit Function
    End If
        
    For intCounter = 1 To intFieldLen
        If Mid(ctlField.Text, intCounter, 1) <> " " Then
            IsBlank = False
            Exit For
        End If
    Next
End Function

'*******************************************************************************
' IsOdd
'
' INPUT:  Whole number.
'
' OUPUT:  n/a
'
' RETURN: True if the number is odd, otherwise false.
'
' NOTES:  n/a
'*******************************************************************************
Public Function IsOdd(ByVal intVal As Integer) As Boolean

    IsOdd = True

    If intVal Mod 2 = 1 Then
        IsOdd = False
    End If

End Function

'******************************************************************************
' IsPositive
'
' INPUT:  A control.
'
' OUTPUT: n/a
'
' RETURN: True if the number is positive, otherwise false.
'
' NOTES:  This function is not in current use.
'******************************************************************************
Public Function IsPositive(ctlField As Control) As Boolean
    Dim RetVal As Integer
    
    IsPositive = False
    
    If CDbl(ctlField.Text) > 0 Then
        IsPositive = True
    End If
End Function

'******************************************************************************
' IsBetween
'
' INPUT:  The number to be evaluate, and the low and high numbers in the range.
'
' OUTPUT: n/a
'
' RETURN: True if the number is between (inclusive), otherwise false.
'
' NOTES:  This function is not in current use.
'******************************************************************************
Public Function IsBetween(ByVal intNumber As Integer, ByVal intLow As Integer, ByVal intHigh As Integer) As Boolean

   If intNumber >= intLow And intNumber <= intHigh Then
      IsBetween = True
   Else
      IsBetween = False
   End If

End Function

'******************************************************************************
' ValidateCurrencyFormat
'
' INPUT:  A control.
'
' OUTPUT: n/a
'
' RETURN: True if it is currency format, oterwise false.
'
' NOTES:  This function is not in current use.
'******************************************************************************
Public Function ValidateCurrencyFormat(ctlField As Control) As Boolean
    Dim RetVal As Integer
    Dim intLength As Integer
    Dim intDecimalPosition As Integer
    
    ValidateCurrencyFormat = True
    
    intLength = Len(ctlField.Text)
    
    intDecimalPosition = InStr(ctlField.Text, ".")

    If (intDecimalPosition + 2) <> intLength Then
        RetVal = MsgBox("Field is not currency format. (i.e., 1.23)", vbExclamation, "Error")
        ctlField.SetFocus
        ValidateCurrencyFormat = False
    End If
End Function

'******************************************************************************
' ValidateNumeric
'
' INPUT:  A control.
'
' OUTPUT: n/a
'
' RETURN: True if the value is a number, otherwise false.
'
' NOTES:  This function is not in current use.
'******************************************************************************
Public Function ValidateNumeric(ctlField As Control) As Boolean
    Dim RetVal As Integer
    
    ValidateNumeric = True
    
    If IsNumeric(ctlField.Text) = False Then
        RetVal = MsgBox("Field is not numeric.", vbExclamation, "Error")
        ctlField.SetFocus
        ValidateNumeric = False
    End If
End Function


