<%
'//////////////////////////////////////////////////////////////////////////////////////////////////
'// File Name: function_global.asp
'// Description: 
'//
'//	Function List : 
'//		Function GenerateKeyCodeGlobal
'//		Function ArrayToStringGlobal
'//		Function PopArrayGlobal
'//		Function PushArrayGlobal
'//		Function BytesToStrGlobal
'//		Function CnvrNullToEmptyGlobal
'//		Function chkEmptyGlobal
'//		Function chkInValuesGlobal
'//		Function chkDisplayStyleGlobal
'//		Function stripHTML
'//		Function ConvertNum
'//		Function FormatDate
'//		Function SetRecordsetToDictionary
'//		Function SetQueryToDictionary
'//		Function FormatDuration
'//		Function FormatAddress
'//		Function FormatNum
'//		Function Exist_QF
'//		Function Exist_QF_Array
'//		Function checkSQLInjection
'//		Function CnvrPhoneNumber
'//		Sub AddTableRow
'//		Sub UpdateTableRow
'//		Sub SetTableField
'//		Function CheckTableField
'//		Sub PrintSelectOptions
'//		Sub PrintSelectCodes
'//		Sub PrintSelectStates
'//		Function PrintHoursMinutesByMinutes
'//		
'// Things to do: 
'//
'//////////////////////////////////////////////////////////////////////////////////////////////////

'/**
'* Generate style attrubute code.
'* @author	jcoh
'* @name	SetBGColor
'* @param	{string} val - A string that is Hex Code RGB
'* @return	{string} HTML style tag
'*/
'================================================================================================================
'= SetBGColor
'= 
'================================================================================================================
Function SetBGColor (ByVal val)
	Dim style : style = ""
	If chkEmptyGlobal(val) = FALSE Then
		style = "style=""background-color:" & val & " !important;"""
	End If 
	Val = ""
	SetBGColor = style
End Function
'/**
'* 
'* @author	jcoh
'* @name	GetHiddenInput
'* @param	{string} name - a name of hidden input.
'* @param	{string} val - value
'* @return	{string} HTML hidden input tag
'*/
'================================================================================================================
'= GetHiddenInput
'= 
'================================================================================================================
Function GetHiddenInput (ByVal name, ByVal val)
	GetHiddenInput = "<input type=""hidden"" name=""" & name & """ value=""" & val & """>"
End Function
'================================================================================================================
'= PrintHoursMinutesByMinutes
'= 
'================================================================================================================
Function PrintHoursMinutesByMinutes (ByVal minutes) 
	Dim str_hours
	Dim str_minutes

	Dim ret_val : ret_val = ""
	If chkEmptyGlobal(minutes) = FALSE Then
		minutes = CDbl(minutes)
		str_hours = CInt(Fix(minutes / 60))
		If str_hours < 10 Then
			str_hours = "0" & str_hours
		End If 

		str_minites = CInt(minutes Mod 60)
		If str_minites < 10 Then
			str_minites = "0" & str_minites
		End If 

		ret_val = str_hours & ":" & str_minites
	End If 

	
	PrintHoursMinutesByMinutes = ret_val

End Function
'================================================================================================================
'= GenerateKeyCode
'= Parameters : key_type ->[2:00],[3:B00],[3a],[3n:B00],[4:100],[4n:A000],[5:BB001],[7:0000000]
'================================================================================================================
Function GenerateKeyCodeGlobal	(_
								ByRef obj_conn, _
								ByVal key_type, _
								ByVal key_id, _
								ByVal key_table _
								)
	Dim sql
	Dim obj_rs
	Dim key_code
	Dim key_prefix
	Dim key_index
	
	If chkEmptyGlobal(key_type) = FALSE AND chkEmptyGlobal(key_id) = FALSE AND chkEmptyGlobal(key_table) = FALSE Then
		Select Case Trim(key_type)
			'--------------------------------------------------------------------------------------------------
			'- [2]
			'--------------------------------------------------------------------------------------------------
			Case "2"
				sql = ""
				sql = sql & " SELECT				TOP 1 "
				sql = sql & "						" & id & " "
				sql = sql & " FROM					" & key_table & " "
				sql = sql & " WHERE					" & id & " BETWEEN '0' AND '99' "
				sql = sql & " ORDER BY				" & id & " DESC "

				Set obj_rs = obj_conn.OpenQuery(sql)

				If obj_rs.EOF Then
					key_code = "00"
				Else
					key_index = CInt(obj_rs(0)) + 1

					If key_index < 10 Then
						key_code = "0" + CStr(key_index)
					Else
						key_code = CStr(key_index)
					End If
				End If

				obj_rs.Close
				Set obj_rs = Nothing
			'--------------------------------------------------------------------------------------------------
			'- [3]
			'--------------------------------------------------------------------------------------------------
			Case "3"
				sql =	""
				sql = sql & " SELECT				TOP 1 "
				sql = sql & "						LEFT(" & key_id & ",1) as col_1 ,"
				sql = sql & "						RIGHT(" & key_id & ",2)as col_2 "
				sql = sql & " FROM					" & key_table & " "
				sql = sql & " WHERE					RIGHT(" & key_id & ",2) BETWEEN '0' AND '99' "
				sql = sql & "						AND LEN(" & key_id & ") = 3 "
				sql = sql & " ORDER BY				" & key_id & " DESC "
				
				Set obj_rs = obj_conn.OpenQuery(sql)

				If obj_rs.EOF Then
					key_code = "B00"
				Else
					key_prefix = obj_rs(0)
					key_index = obj_rs(1)
					

					If key_prefix = "9" OR key_prefix = "0" OR chkEmptyGlobal(key_prefix) = TRUE Then
						key_prefix = "B"
					End If

					If chkEmptyGlobal(key_index) = TRUE Then
						key_index = 1
					Else
						key_index = CInt(key_index) + 1
					End If

					if key_index < 10 	then
						key_code = key_prefix + "0" + Cstr(key_index)
					elseif key_index < 100 then	
						key_code = key_prefix + Cstr(key_index)
					elseif key_index = 100 and key_prefix = "B" then
						key_code = "C01"
					elseif key_index = 100 and key_prefix = "C" then
						key_code = "D01"
					elseif key_index = 100 and key_prefix = "D" then
						key_code = "E01"
					elseif key_index = 100 and key_prefix = "E" then
						key_code = "F01"
					elseif key_index = 100 and key_prefix = "F" then
						key_code = "G01"
					elseif key_index = 100 and key_prefix = "G" then
						key_code = "H01"
					elseif key_index = 100 and key_prefix= "H" then
						key_code = "I01"
					elseif key_index = 100 and key_prefix = "I" then
						key_code = "J01"
					elseif key_index = 100 and key_prefix = "J" then
						key_code = "K01"
					elseif key_index = 100 and key_prefix = "K" then
						key_code = "L01"
					elseif key_index = 100 and key_prefix = "L" then
						key_code = "M01"
					elseif key_index = 100 and key_prefix = "M" then
						key_code = "N01"
					elseif key_index = 100 and key_prefix = "N" then
						key_code = "O01"
					elseif key_index = 100 and key_prefix = "O" then
						key_code = "P01"
					elseif key_index = 100 and key_prefix = "P" then
						key_code = "Q01"
					elseif key_index = 100 and key_prefix = "Q" then
						key_code = "R01"
					elseif key_index = 100 and key_prefix = "R" then
						key_code = "S01"
					elseif key_index = 100 and key_prefix = "S" then
						key_code = "T01"
					elseif key_index = 100 and key_prefix = "T" then
						key_code = "U01"
					elseif key_index = 100 and key_prefix = "U" then
						key_code = "V01"
					elseif key_index = 100 and key_prefix = "V" then
						key_code = "W01"
					elseif key_index = 100 and key_prefix = "W" then
						key_code = "X01"
					elseif key_index = 100 and key_prefix = "X" then
						key_code = "Y01"
					elseif key_index = 100 and key_prefix = "Y" then
						key_code = "Z01"
					end if
				End If

				obj_rs.Close
				Set obj_rs = Nothing
			'--------------------------------------------------------------------------------------------------
			'- [3a]
			'--------------------------------------------------------------------------------------------------
			Case "3a"				
				sql = ""
				sql = sql & " SELECT				top 1 "
				sql = sql & "						" & key_id & " as id "
				sql = sql & " FROM					" & key_table & " "
				sql = sql & " ORDER BY				" & key_id & " DESC"

				Set obj_rs = obj_conn.OpenQuery(sql)

				If obj_rs.EOF Then
					key_code = "001"
				Else
					key_code = obj_rs(0)
					
					If IsNumeric(key_code) = TRUE Then
						key_index = CDbl(key_code) + 1

						If key_index <= 999 Then
							if key_index < 10 	then
								key_code = "00" + CStr(key_index)
							elseif key_index < 100 then	
								key_code =  "0" + CStr(key_index)
							Else
								key_code = CStr(key_index)
							End If 
						Else
							key_code = ""
						End If
					End If

					If IsNumeric(key_code) = FALSE Then
						
						If chkEmptyGlobal(key_code) = TRUE Then
							key_code = "A01"
						Else
							key_prefix	= Left(key_code,1)
							key_index	= Right(key_code,2)
							key_index = CInt(key_index) + 1
						End If 

						if key_index < 10 	then
							key_code = key_prefix + "0" + Cstr(key_index)
						elseif key_index < 100 then	
							key_code = key_prefix + Cstr(key_index)
						elseif key_index = 100 and key_prefix = "A" then
							key_code = "B01"
						elseif key_index = 100 and key_prefix = "B" then
							key_code = "C01"
						elseif key_index = 100 and key_prefix = "C" then
							key_code = "D01"
						elseif key_index = 100 and key_prefix = "D" then
							key_code = "E01"
						elseif key_index = 100 and key_prefix = "E" then
							key_code = "F01"
						elseif key_index = 100 and key_prefix = "F" then
							key_code = "G01"
						elseif key_index = 100 and key_prefix = "G" then
							key_code = "H01"
						elseif key_index = 100 and key_prefix = "H" then
							key_code = "I01"
						elseif key_index = 100 and key_prefix = "I" then
							key_code = "J01"
						elseif key_index = 100 and key_prefix = "J" then
							key_code = "K01"
						elseif key_index = 100 and key_prefix = "K" then
							key_code = "L01"
						elseif key_index = 100 and key_prefix = "L" then
							key_code = "M01"
						elseif key_index = 100 and key_prefix = "M" then
							key_code = "N01"
						elseif key_index = 100 and key_prefix = "N" then
							key_code = "O01"
						elseif key_index = 100 and key_prefix = "O" then
							key_code = "P01"
						elseif key_index = 100 and key_prefix = "P" then
							key_code = "Q01"
						elseif key_index = 100 and key_prefix = "Q" then
							key_code = "R01"
						elseif key_index = 100 and key_prefix = "R" then
							key_code = "S01"
						elseif key_index = 100 and key_prefix = "S" then
							key_code = "T01"
						elseif key_index = 100 and key_prefix = "T" then
							key_code = "U01"
						elseif key_index = 100 and key_prefix = "U" then
							key_code = "V01"
						elseif key_index = 100 and key_prefix = "V" then
							key_code = "W01"
						elseif key_index = 100 and key_prefix = "W" then
							key_code = "X01"
						elseif key_index = 100 and key_prefix = "X" then
							key_code = "Y01"
						elseif key_index = 100 and key_prefix = "Y" then
							key_code = "Z01"
						end if

					End If ' IsNumeric(key_code) = FALSE Then
				End If ' obj_rs.EOF Then
				obj_rs.Close
				Set obj_rs = Nothing
			'--------------------------------------------------------------------------------------------------
			'- [3n]
			'--------------------------------------------------------------------------------------------------
			Case "3n"
				sql = ""
				sql = sql & " SELECT				top 1 "
				sql = sql & "						LEFT( " & key_id & ",1) as col_1 ,"
				sql = sql & "						RIGHT(" & key_id & ",2) as col_2 "
				sql = sql & " FROM					" & key_table & " "
				sql = sql & " WHERE					RIGHT(" & key_id & ",2) BETWEEN '10' AND '99' "
				sql = sql & " ORDER BY				" & key_id & " DESC"
				
				Set obj_rs = obj_conn.OpenQuery(sql)

				If obj_rs.EOF Then
					key_code = "000"
				Else
					key_prefix	= obj_rs(0)
					key_index	= obj_rs(1)

					If key_prefix = "9" OR chkEmptyGlobal(key_prefix) = TRUE Then
						key_prefix = "B"
					End If

					If chkEmptyGlobal(key_index) = TRUE Then
						key_index = 1
					Else
						key_index = CInt(key_index) + 1
					End If
					
					if key_index < 10 	then
						key_code = key_prefix + "0" + Cstr(key_index)
					elseif key_index < 100 then	
						key_code = key_prefix + Cstr(key_index)
					elseif key_index = 100 and key_prefix = "B" then
						key_code = "C01"
					elseif key_index = 100 and key_prefix = "C" then
						key_code = "D01"
					elseif key_index = 100 and key_prefix = "D" then
						key_code = "E01"
					elseif key_index = 100 and key_prefix = "E" then
						key_code = "F01"
					elseif key_index = 100 and key_prefix = "F" then
						key_code = "G01"
					elseif key_index = 100 and key_prefix = "G" then
						key_code = "H01"
					elseif key_index = 100 and key_prefix = "H" then
						key_code = "I01"
					elseif key_index = 100 and key_prefix = "I" then
						key_code = "J01"
					elseif key_index = 100 and key_prefix = "J" then
						key_code = "K01"
					elseif key_index = 100 and key_prefix = "K" then
						key_code = "L01"
					elseif key_index = 100 and key_prefix = "L" then
						key_code = "M01"
					elseif key_index = 100 and key_prefix = "M" then
						key_code = "N01"
					elseif key_index = 100 and key_prefix = "N" then
						key_code = "O01"
					elseif key_index = 100 and key_prefix = "O" then
						key_code = "P01"
					elseif key_index = 100 and key_prefix = "P" then
						key_code = "Q01"
					elseif key_index = 100 and key_prefix = "Q" then
						key_code = "R01"
					elseif key_index = 100 and key_prefix = "R" then
						key_code = "S01"
					elseif key_index = 100 and key_prefix = "S" then
						key_code = "T01"
					elseif key_index = 100 and key_prefix = "T" then
						key_code = "U01"
					elseif key_index = 100 and key_prefix = "U" then
						key_code = "V01"
					elseif key_index = 100 and key_prefix = "V" then
						key_code = "W01"
					elseif key_index = 100 and key_prefix = "W" then
						key_code = "X01"
					elseif key_index = 100 and key_prefix = "X" then
						key_code = "Y01"
					elseif key_index = 100 and key_prefix = "Y" then
						key_code = "Z01"
					end if
				End If
				obj_rs.Close
				Set obj_rs = Nothing
			'--------------------------------------------------------------------------------------------------
			'- [4]
			'--------------------------------------------------------------------------------------------------
			Case "4"
				sql = ""
				sql = sql & " SELECT				TOP 1 "
				sql = sql & "						" & key_id & " "
				sql = sql & " FROM					" & key_table & " "
				sql = sql & " WHERE					" & key_id & " BETWEEN '100' AND '999' "
				sql = sql & " ORDER BY				" & key_id & " DESC"

				Set obj_rs = obj_conn.OpenQuery(sql)
				If obj_rs.EOF Then
					key_code = "100"
				Else
					key_index	= CInt(obj_rs(0)) + 1
					if key_index < 10 		then
						key_code = "00" + CStr(key_index)
					elseif key_index < 100 	then
						key_code = "0" + CStr(key_index)
					else
						key_code = CStr(key_index)
					end if
				End If
				obj_rs.Close
				Set obj_rs = Nothing
			'--------------------------------------------------------------------------------------------------
			'- [4n]
			'--------------------------------------------------------------------------------------------------
			Case "4n"
				sql = ""
				sql = sql & " SELECT				TOP 1 "
				sql = sql & "						LEFT(" & key_id & ",1) as col_1 ,"
				sql = sql & "						RIGHT(" & key_id & ",3) as col_2 "
				sql = sql & " FROM					" & key_table & " "
				sql = sql & " WHERE					RIGHT(" & key_id & ",3) BETWEEN '0' AND '999' "
				sql = sql & " ORDER BY				" & key_id & " DESC "

				Set obj_rs = obj_conn.OpenQuery(sql)

				If obj_rs.EOF Then
					key_code = "A000"
				Else
					key_prefix	= obj_rs(0)
					key_index	= obj_rs(1)

					If key_prefix = "9" OR chkEmptyGlobal(key_prefix) = TRUE Then
						key_prefix = "A"
					End If

					If chkEmptyGlobal(key_index) = TRUE Then
						key_index = 1
					Else
						key_index = CInt(key_index) + 1
					End If

					if key_index < 10 	then
						key_code	= key_prefix + "00" + CStr(key_index)
					elseif key_index < 100 then
						key_code	=	key_prefix + "0" + CStr(key_index)
					elseif key_index < 1000 then	
						key_code = key_prefix + CStr(key_index)
					elseif key_index = 1000 and key_prefix = "A" then
						key_code = "B001"
					elseif key_index = 1000 and key_prefix = "B" then
						key_code = "C001"
					elseif key_index = 1000 and key_prefix = "C" then
						key_code = "D001"
					elseif key_index = 1000 and key_prefix = "D" then
						key_code = "E001"
					elseif key_index = 1000 and key_prefix = "E" then
						key_code = "F001"
					elseif key_index = 1000 and key_prefix = "F" then
						key_code = "G001"
					elseif key_index = 1000 and key_prefix = "G" then
						key_code = "H001"
					elseif key_index = 1000 and key_prefix = "H" then
						key_code = "I001"
					elseif key_index = 1000 and key_prefix = "I" then
						key_code = "J001"
					elseif key_index = 1000 and key_prefix = "J" then
						key_code = "K001"
					elseif key_index = 1000 and key_prefix = "K" then
						key_code = "L001"
					elseif key_index = 1000 and key_prefix = "L" then
						key_code = "M001"
					elseif key_index = 1000 and key_prefix = "M" then
						key_code = "N001"
					elseif key_index = 1000 and key_prefix = "N" then
						key_code = "O001"
					elseif key_index = 1000 and key_prefix = "O" then
						key_code = "P001"
					elseif key_index = 1000 and key_prefix = "P" then
						key_code = "Q001"
					elseif key_index = 1000 and key_prefix = "Q" then
						key_code = "R001"
					elseif key_index = 1000 and key_prefix = "R" then
						key_code = "S001"
					elseif key_index = 1000 and key_prefix = "S" then
						key_code = "T001"
					elseif key_index = 1000 and key_prefix = "T" then
						key_code = "U001"
					elseif key_index = 1000 and key_prefix = "U" then
						key_code = "V001"
					elseif key_index = 1000 and key_prefix = "V" then
						key_code = "W001"
					elseif key_index = 1000 and key_prefix = "W" then
						key_code = "X001"
					elseif key_index = 1000 and key_prefix = "X" then
						key_code = "Y001"
					elseif key_index = 1000 and key_prefix = "Y" then
						key_code = "Z001"
					end if
				End If
				obj_rs.Close
				Set obj_rs = Nothing

			'--------------------------------------------------------------------------------------------------
			'- [5]
			'--------------------------------------------------------------------------------------------------
			Case "5"	
				sql =	""
				sql = sql & " SELECT				TOP 1 "
				sql = sql & "						LEFT(" & key_id & ",2) as col_1 ,"
				sql = sql & "						RIGHT(" & key_id & ",3) as col_2 "
				sql = sql & " FROM					" & key_table & " "
				sql = sql & " WHERE					LEFT(" & key_id & ",2) > 'AZ' "
				sql = sql & "						and RIGHT(" & key_id & ",3) BETWEEN '0' AND '999' "
				sql = sql & " ORDER BY				" & key_id & " DESC "
				
				Set obj_rs = obj_conn.OpenQuery(sql)

				If obj_rs.EOF Then
					key_code = "BB001"
				Else
					key_prefix	= obj_rs(0)
					key_index	= obj_rs(1)

					If key_prefix = "99" OR chkEmptyGlobal(key_prefix) = TRUE Then
						key_prefix = "BB"
					End If

					If chkEmptyGlobal(key_index) = TRUE Then
						key_index = 1
					Else
						key_index = CInt(key_index) + 1
					End If

					if key_index < 10 	then
						key_code = key_prefix + "00" + Cstr(key_index)
					elseif key_index < 100 then	
						key_code = key_prefix + "0" + Cstr(key_index)
					elseif key_index < 1000 then
						key_code = key_prefix + Cstr(key_index)
					elseif key_index = 1000 and key_prefix = "BB" then
						key_code = "CC001"
					elseif key_index = 1000 and key_prefix = "CC" then
						key_code = "DD001"
					elseif key_index = 1000 and key_prefix = "DD" then
						key_code = "EE001"
					elseif key_index = 1000 and key_prefix = "EE" then
						key_code = "FF001"
					elseif key_index = 1000 and key_prefix = "FF" then
						key_code = "GG001"
					elseif key_index = 1000 and key_prefix = "GG" then
						key_code = "HH001"
					elseif key_index = 1000 and key_prefix = "HH" then
						key_code = "II001"
					elseif key_index = 1000 and key_prefix = "II" then
						key_code = "JJ001"
					elseif key_index = 1000 and key_prefix = "JJ" then
						key_code = "KK001"
					elseif key_index = 1000 and key_prefix = "KK" then
						key_code = "LL001"
					elseif key_index = 1000 and key_prefix = "LL" then
						key_code = "MM001"
					elseif key_index = 1000 and key_prefix = "MM" then
						key_code = "NN001"
					elseif key_index = 1000 and key_prefix = "NN" then
						key_code = "OO001"
					elseif key_index = 1000 and key_prefix = "OO" then
						key_code = "PP001"
					elseif key_index = 1000 and key_prefix = "PP" then
						key_code = "QQ001"
					elseif key_index = 1000 and key_prefix = "QQ" then
						key_code = "RR001"
					elseif key_index = 1000 and key_prefix = "RR" then
						key_code = "SS001"
					elseif key_index = 1000 and key_prefix = "SS" then
						key_code = "TT001"
					elseif key_index = 1000 and key_prefix = "TT" then
						key_code = "UU001"
					elseif key_index = 1000 and key_prefix = "UU" then
						key_code = "VV001"
					elseif key_index = 1000 and key_prefix = "VV" then
						key_code = "WW001"
					elseif key_index = 1000 and key_prefix = "WW" then
						key_code = "XX001"
					elseif key_index = 1000 and key_prefix = "XX" then
						key_code = "YY001"
					elseif key_index = 1000 and key_prefix = "YY" then
						key_code = "ZZ001"
					end if
				End If
				obj_rs.Close
				Set obj_rs = Nothing

			'--------------------------------------------------------------------------------------------------
			'- [7]
			'--------------------------------------------------------------------------------------------------
			Case "7"	
				sql = ""
				sql = sql & " SELECT				TOP 1 "
				sql = sql & "						" & key_id & " "
				sql = sql & " FROM					" & key_table & " "
				sql = sql & " WHERE					" & key_id & " BETWEEN '0000000' AND '9999999' "
				sql = sql & " ORDER BY				" & key_id &" DESC"
				Set obj_rs = obj_conn.OpenQuery(sql)

				If obj_rs.EOF Then
					key_code = "0000000"
				Else
		  			key_index = CDbl(obj_rs(0)) + 1
					if key_index < 10 then
						key_code = "000000" + Cstr(key_index)
					elseif key_index < 100 then
						key_code = "00000" 	+ Cstr(key_index)
					elseif key_index < 1000 then
						key_code = "0000" + Cstr(key_index)
					elseif key_index < 10000 then
						key_code = "000" + Cstr(key_index)
					elseif key_index < 100000 then
						key_code = "00" + Cstr(key_index)
					elseif key_index < 1000000 then
						key_code = "0" + Cstr(key_index)
					else
						key_code = Cstr(key_index)
					end if
				End If
				obj_rs.Close
				Set obj_rs = Nothing
		End Select
	End If
	
	GenerateKeyCodeGlobal = key_code
End Function
'================================================================================================================
'= ArrayToStringGlobal
'================================================================================================================
Function ArrayToStringGlobal(ByVal arr_data,splitter)
	Dim str_val
	Dim ret_val : ret_val = ""
	Dim i
	Dim len_arr
	
	If IsArray(arr_data) = TRUE AND Trim(splitter) <> "" AND Len(splitter) = 1 Then
		len_arr = UBound(arr_data)
		If len_arr <> -1 Then
			For i = 0 To len_arr
				If i = len_arr Then
					ret_val = ret_val & arr_data(i)
				Else
					ret_val = ret_val & arr_data(i) & splitter
				End If
			Next
		End If
	End If
	ArrayToStringGlobal = ret_val
End Function
'================================================================================================================
'= PushArrayGlobal
'================================================================================================================
Function PushArrayGlobal (ByRef arr_data, ByVal val_data)
	Dim arr_temp
	
	If IsArray(arr_data) = TRUE Then
		arr_temp = arr_data
		
		ReDim Preserve arr_temp(UBound(arr_temp)+1)
		arr_temp(UBound(arr_temp)) = val_data
		arr_data = arr_temp
	End If

	PushArrayGlobal = arr_data
End Function

'================================================================================================================
'= PopArrayGlobal
'================================================================================================================
Function PopArrayGlobal (ByRef arr_data)
	Dim arr_temp
	Dim ret_val
	Dim i
	
	If IsArray(arr_data) = TRUE Then
		If UBound(arr_data) <> -1 Then
			ret_val = arr_data(0)
			arr_temp = arr_data
			ReDim Preserve arr_data(UBound(arr_data)-1)

			For i = 0 To UBound(arr_data)
				arr_data(i) = arr_temp(i+1)
			Next
		End If
	End If

	PopArrayGlobal = ret_val
End Function
'================================================================================================================
'= BytesToStrGlobal
'================================================================================================================
Function BytesToStrGlobal (ByVal bytes) 
    Dim Stream 
    Set Stream = Server.CreateObject("Adodb.Stream") 
    Stream.Type = 1 
    'adTypeBinary 
    Stream.Open 
    Stream.Write bytes 
    Stream.Position = 0 
    Stream.Type = 2 'adTypeText 
    Stream.Charset = "utf-8" 
    BytesToStrGlobal = Stream.ReadText 
    Stream.Close 
    Set Stream = Nothing 
End Function

'================================================================================================================
'= URLDecodeGlobal
'================================================================================================================
Function URLDecodeGlobal(ByVal sConvert)
    Dim aSplit
    Dim sOutput
    Dim I
    If IsNull(sConvert) Then
       URLDecode = ""
       Exit Function
    End If

    ' convert all pluses to spaces
    sOutput = REPLACE(sConvert, "+", " ")

    ' next convert %hexdigits to the character
    aSplit = Split(sOutput, "%")

    If IsArray(aSplit) Then
      sOutput = aSplit(0)
      For I = 0 to UBound(aSplit) - 1
        sOutput = sOutput & _
          Chr("&H" & Left(aSplit(i + 1), 2)) &_
          Right(aSplit(i + 1), Len(aSplit(i + 1)) - 2)
      Next
    End If

    URLDecodeGlobal = sOutput
End Function


'================================================================================================================
'= PrintSelectCodes
'================================================================================================================
Sub PrintSelectCodes (ByRef obj_conn, ByVal conn_type,code_id,code)
	Call PrintSelectOptions	(_
							obj_conn, _
							conn_type, _
							"SELECT code, code_name FROM tbCode WHERE code_id='" & code_id & "' AND ISNULL(code_name,'') <> '' ORDER BY display_order", _
							code _
							)
End Sub
'================================================================================================================
'= PrintSelectStates
'================================================================================================================
Sub PrintSelectStates (ByRef obj_conn, ByVal conn_type,state_code)
	Call PrintSelectOptions	(_
							obj_conn, _
							conn_type, _
							"SELECT abbr, state FROM states WHERE abbr<>'ZZ' ORDER BY abbr", _
							state_code _
							)
End Sub
'================================================================================================================
'= PrintSelectOptions
'================================================================================================================
Sub PrintSelectOptions	(_
						ByRef obj_conn, _
						conn_type, _
						query, _
						code _
						)
	
	Dim adoRs
	Dim id
	Dim id_name
	if query = "" then
		exit sub
	end If

	If UCase(conn_type) = "DBCON" Then
		Set adoRs = Server.CreateObject("ADODB.Recordset")
		adoRs.Open query, obj_conn, 1,3
	Else
		Set adoRs = obj_conn.OpenQuery(query)
	End If

	if Not adoRs.EOF Then
		Do Until adoRs.EOF
			id = adoRs(0)
			id_name = adoRs(1)
			if isNull(id) then id = ""
			if isNull(id_name) then id_name = ""
			if isNull(code) then code = ""
			
			if cstr(code) = cstr(id) then
				str = str&"<option value="""&id&""" selected>"&UCase(Trim(id_name))&"</option>"
			else
				str = str&"<option value="""&id&""">"&UCase(Trim(id_name))&"</option>"
			end if


			adoRs.MoveNext
		Loop
	end if
	Set adoRs = Nothing

	Response.Write str
End Sub
'================================================================================================================
'= CnvrNullToEmptyGlobal
'================================================================================================================
Function CnvrNullToEmptyGlobal(ByVal str)
	Dim ret_val 
	If chkEmptyGlobal(str) = TRUE Then
		ret_val = ""
	Else
		ret_val = str
	End If
	CnvrNullToEmptyGlobal = ret_val
End Function

'================================================================================================================
'= CnvrEmptyToDefaultGlobal
'================================================================================================================
Function CnvrEmptyToDefaultGlobal(ByVal str_val,ByVal str_def)
	Dim ret_val 
	If chkEmptyGlobal(str_val) = TRUE AND chkEmptyGlobal(str_def) = FALSE Then
		ret_val = str_def
	Else
		ret_val = str_val
	End If
	CnvrEmptyToDefaultGlobal = ret_val
End Function

'==================================================================================================
'= chkEmptyGlobal
'==================================================================================================
Function chkEmptyGlobal (Val)
	Dim RetVal 
	Val = Trim(Val)

	If IsEmpty(Val) OR IsNull(Val) OR Val = "" Then
		RetVal = TRUE
	Else
		RetVal = FALSE
	End If

	chkEmptyGlobal = RetVal 
End Function

'==================================================================================================
'= chkInValuesGlobal
'==================================================================================================
Function chkInValuesGlobal (ByVal strVal,ByVal strVals)
	Dim arrVal
	Dim itmVal
	Dim RetVal : RetVal = FALSE

	arrVal = Split(strVals,"|")

	For Each itmVal in arrVal
		If strVal = itmVal Then
			RetVal = TRUE
		End If 
	Next

	chkInValuesGlobal = RetVal
End Function

'==================================================================================================
'= chkDisplayStyleGlobal 
'==================================================================================================
Function chkDisplayStyleGlobal (ByVal Val)
	Dim RetVal 
	If chkEmptyGlobal (Val) = TRUE Then
		RetVal = "style=""display:none;"""
	End If
	chkDisplayStyleGlobal = RetVal
End Function


'==================================================================================================
'= stripHTML 
'==================================================================================================
Function stripHTML(ByVal strHTML)
  Dim objRegExp, strOutput, tempStr
  Set objRegExp = New Regexp
  objRegExp.IgnoreCase = True
  objRegExp.Global = True
  objRegExp.Pattern = "<(.|n)+?>"

  If IsNull(strHTML) = FALSE Then
	  'Replace all HTML tag matches with the empty string
	  strOutput = objRegExp.Replace(strHTML, "")
	  'Replace all < and > with &lt; and &gt;
	  strOutput = Replace(strOutput, "<", "&lt;")
	  strOutput = Replace(strOutput, ">", "&gt;")
	  stripHTML = strOutput    'Return the value of strOutput
	  Set objRegExp = Nothing
  Else
	stripHTML = strHTML
  End If
End Function
'==================================================================================================
'= Name: ConvertNum
'= Parameter : [CINT],[CDBL]
'= Desc: 0 or Number
'==================================================================================================
Function ConvertNum		(_
						ByVal data, _
						ByVal data_type _
						)
	Dim ret_val :  ret_val = 0

	If chkEmptyGlobal(data) = FALSE AND chkEmptyGlobal(data_type) = FALSE AND IsNumeric(Trim(data)) = TRUE Then
		Select Case UCase(Trim(data_type))
			Case "CINT"
				ret_val = CInt(data)
			Case "CDBL"
				ret_val = CDbl(data)
		End Select
	End If

	ConvertNum = ret_val
End Function

'==================================================================================================
'= Name: FormatDate
'= Desc: 
'==================================================================================================
Function FormatDate		(_
						ByVal datetime, _
						ByVal date_type _
						)
	Dim ret_val : ret_val = ""


	If chkEmptyGlobal(datetime) = FALSE AND IsDate(datetime) = TRUE Then
		
		Select Case LCase(Trim(date_type))
			Case "0","101","mm/dd/yyyy" ' 03/30/2017
				ret_val = Right("00"&month(datetime),2)&"/"&Right("00"&day(datetime),2)&"/"&Right("0000"&year(datetime),4)
			Case "1","mm/dd/yyyy hh:mm:ss" ' 03/30/2017 15:06:04
				ret_val = Right("00"&month(datetime),2)&"/"&Right("00"&day(datetime),2)&"/"&Right("0000"&year(datetime),4)&" "&Right("00"&hour(datetime),2)&":"&Right("00"&minute(datetime),2)&":"&Right("00"&second(datetime),2)	
			Case "2","month_name dd, yyyy" ' March 30, 2017
				ret_val = monthname(month(datetime))&" "&Right("00"&day(datetime),2)&", "&Right("0000"&year(datetime),4)
			Case "3","month_name dd, yyyy hh:mm:ss" ' March 30, 2017 15:13:48
				ret_val = monthname(month(datetime))&" "&Right("00"&day(datetime),2)&", "&Right("0000"&year(datetime),4)&" "&Right("00"&hour(datetime),2)&":"&Right("00"&minute(datetime),2)&":"&Right("00"&second(datetime),2)
			Case "4","102","yyyy.mm.dd" ' 2017.03.30
				ret_val = Right("0000"&year(datetime),4)&"."&Right("00"&month(datetime),2)&"."&Right("00"&day(datetime),2)
			Case "5","mm/dd" ' 03/30
				ret_val = Right("00"&month(datetime),2)&"/"&Right("00"&day(datetime),2)
			Case "6","month_name, yyyy" ' March, 2017
				ret_val = monthname(month(datetime))&", "&Right("0000"&year(datetime),4)
			Case "7","mm/dd/yy" ' 03/30/17
				ret_val = Right("00"&month(datetime),2)&"/"&Right("00"&day(datetime),2)&"/"&Right("0000"&year(datetime),2)
			Case "8","41","yyyy-mm-dd" ' 2017-03-30
				ret_val = Right("0000"&year(datetime),4)&"-"&Right("00"&month(datetime),2)&"-"&Right("00"&day(datetime),2)
			Case "9","31","yyyymmdd" ' 20170330
				ret_val = Right("0000"&year(datetime),4)&Right("00"&month(datetime),2)&Right("00"&day(datetime),2)
			Case "10","mm/dd/yy hh:mm" ' 03/30/17 10:10
				ret_val = Right("00"&month(datetime),2)&"/"&Right("00"&day(datetime),2)&"/"&Right("0000"&year(datetime),2) & " " & Right("00"&hour(datetime),2)&":"&Right("00"&minute(datetime),2)
			Case "11","mm/yyyy" ' 03/2017
				ret_val = Right("00"&month(datetime),2)&"/"&Right("0000"&year(datetime),4)
			Case "12","mm/yy" ' 03/17
				ret_val = Right("00"&month(datetime),2)&"/"&Right("00"&day(datetime),2)
			Case "20","mmddyyyy" ' 03302017
				ret_val = Right("00"&month(datetime),2)&Right("00"&day(datetime),2)&Right("0000"&year(datetime),4)
			Case "21","mmddyyyyhhmmss" ' 03302017152451
				ret_val = Right("00"&month(datetime),2)&Right("00"&day(datetime),2)&Right("0000"&year(datetime),4)&Right("00"&hour(datetime),2)&Right("00"&minute(datetime),2)&Right("00"&second(datetime),2)
			Case "30","yyyy/mm/dd" ' 2017/03/30
				ret_val = Right("0000"&year(datetime),4)&"/"&Right("00"&month(datetime),2)&"/"&Right("00"&day(datetime),2)

			Case "35","yyyymmddhhmmss" ' 20170330152818
				ret_val = Right("0000"&year(datetime),4)&Right("0000"&month(datetime),2)&Right("0000"&day(datetime),2)&Right("0000"&hour(datetime),2)&Right("0000"&minute(datetime),2)&Right("0000"&second(datetime),2)
			Case "40","yyyy-mm-ddThh:mm:ss" ' 2017-03-30T15:29:30
				ret_val = Right("0000"&year(datetime),4)&"-"&Right("00"&month(datetime),2)&"-"&Right("00"&day(datetime),2)&"T"&Right("00"&hour(datetime),2)&":"&Right("00"&minute(datetime),2)&":"&Right("00"&second(datetime),2)
		End Select
	End If

	FormatDate = ret_val
End Function

'==================================================================================================
'= Name: SetRecordsetToDictionary
'= Desc: Dictionary Object
'==================================================================================================
Function SetRecordsetToDictionary (ByRef obj_rs)
	Dim field
	Dim obj_dic

	Set obj_dic = Server.CreateObject("Scripting.Dictionary")

	If IsObject(obj_rs) = TRUE Then
		If Not obj_rs.EOF Then
			For Each field In obj_rs.Fields
				If IsObject(field) = TRUE Then
					obj_dic.Add field.Name,field.Value
				End If
			Next
		End If
	End If

	Set SetRecordsetToDictionary = obj_dic

End Function


'==================================================================================================
'= Name: SetQueryToDictionary
'= Desc: Dictionary Object
'==================================================================================================
Function SetQueryToDictionary	(_
								ByRef obj_conn, _
								ByVal conn_type, _
								ByVal SQL _
								)

	Dim field
	Dim obj_dic
	Dim obj_rs

	Set obj_dic = Server.CreateObject("Scripting.Dictionary")

	If chkEmptyGlobal(SQL) = FALSE Then

		If UCase(conn_type) = "DBCON" Then
			Set obj_rs = Server.CreateObject("ADODB.Recordset")
			obj_rs.Open SQL, obj_conn, 1,3
		Else
			Set obj_rs = obj_conn.OpenQuery(SQL)
		End If

		If IsObject(obj_rs) = TRUE Then
			If Not obj_rs.EOF Then
				For Each field In obj_rs.Fields
					If IsObject(field) = TRUE Then
						obj_dic.Add field.Name,field.Value
					End If
				Next
			End If
		End If
	End If

	obj_rs.Close
	Set obj_rs = Nothing
	Set SetQueryToDictionary = obj_dic

End Function
'==================================================================================================
'= Name: FormatDuration
'= Desc: Year(s) Month(s)
'==================================================================================================
Function FormatDuration (_
						ByVal num_year,_
						ByVal num_month _
						)
	Dim ret_val : ret_val = ""

	If chkEmptyGlobal(num_year) = FALSE Then
		If CInt(num_year) = 1 Then
			ret_val = ret_val & num_year & " Year " 
		Else
			ret_val = ret_val & num_year & " Years " 
		End If
	End If


	If chkEmptyGlobal(num_month) = FALSE Then	
	
		If CInt(num_month) = 1 Then
			ret_val = ret_val & num_month & " Month" 
		Else
			ret_val = ret_val & num_month & " Months" 
		End If
	End If
	
	FormatDuration = ret_val
End Function
'==================================================================================================
'= Name: FormatAddress
'= Desc: Full Address
'==================================================================================================
Function FormatAddress	(_
						ByVal street, _
						ByVal street_add, _
						ByVal city, _
						ByVal state, _
						ByVal zip _
						)
	Dim ret_val : ret_val = ""

	If chkEmptyGlobal(street) = FALSE Then
		ret_val = ret_val & street
	End If 
	
	If chkEmptyGlobal(street_add) = FALSE Then
		ret_val = ret_val & " " & street_add
	End If 

	If chkEmptyGlobal(street) = FALSE OR chkEmptyGlobal(street_add) = FALSE Then
		ret_val = ret_val & ", "
	End If

	If chkEmptyGlobal(city) = FALSE Then
		ret_val = ret_val & city & ", "
	End If

	If chkEmptyGlobal(state) = FALSE Then
		ret_val = ret_val & state & " "
	End If
	
	If chkEmptyGlobal(zip) = FALSE Then
		ret_val = ret_val & zip
	End If

	FormatAddress = ret_val
End Function
'==================================================================================================
'= Name: FormatNum
'= Desc: 
'==================================================================================================
Function FormatNum	(ByVal src,ByVal i)
	Dim ret_val

	If chkEmptyGlobal(src) = FALSE AND IsNumeric(trim(src)) = TRUE Then
		ret_val = FormatNumber(trim(src),i)
	Else
		ret_val = ""
	End If
	FormatNum = ret_val

End Function

'==================================================================================================
'= Name: Exist_QF
'= Desc: TRUE/FALSE
'==================================================================================================
Function Exist_QF (ByVal v_name, ByVal v_type) 
	Dim ret_val : ret_val = FALSE
	
	If Trim(v_name) <> "" Then

		Select Case v_type
			Case ""
				If Request(v_name).Count > 0 Then
					ret_val = TRUE
				End If
			Case "f"
				If Request.Form(v_name).Count > 0 Then
					ret_val = TRUE
				End If
			Case "q"
				If Request.QueryString(v_name).Count > 0 Then
					ret_val = TRUE
				End If
		End Select

	End If

	Exist_QF = ret_val
End Function
'==================================================================================================
'= Name: Exist_QF_Array
'= Desc: TRUE/FALSE
'==================================================================================================
Function Exist_QF_Array (ByVal v_name, ByVal v_type, ByVal i)
	Dim ret_val : ret_val = FALSE
	
	If Trim(v_name) <> "" Then

		Select Case v_type
			Case ""
				If Request(v_name).Count >= i Then
					ret_val = TRUE
				End If
			Case "f"
				If Request.Form(v_name).Count >= i Then
					ret_val = TRUE
				End If
			Case "q"
				If Request.QueryString(v_name).Count >= i Then
					ret_val = TRUE
				End If
		End Select
	End If
	
	Exist_QF_Array = ret_val
End Function

'==================================================================================================
'= Name: AddTableRow
'= Desc: 
'==================================================================================================
Sub AddTableRow	(_
				ByRef obj_rs _
				)
	obj_rs.AddNew
End Sub
'==================================================================================================
'= Name: UpdateTableRow
'= Desc: 
'==================================================================================================
Sub UpdateTableRow (_
				ByRef obj_rs _
				)
	If CheckTableField(obj_rs,"date_edit") = TRUE Then
		obj_rs.Fields("date_edit") = Now()
	End If
	obj_rs.Update
End Sub

'==================================================================================================
'= Name: CheckTableField
'= Desc: 
'==================================================================================================
Function CheckTableField	(_
							ByRef obj_rs, _
							ByVal field_name _
							)
	Dim ret_val : ret_val = FALSE
	Dim field
	
	If IsObject(obj_rs) = TRUE Then
		For Each field In obj_rs.Fields
			If Trim(LCase(field.Name)) = Trim(LCase(field_name)) Then
				'Response.Write field.Name & ":" & field_name & "<BR>"
				ret_val = TRUE
				Exit For
			End If
		Next
	End If
	
	CheckTableField = ret_val
End Function


'==================================================================================================
'= Name: SetDataField
'= Desc: 
'==================================================================================================
Sub SetDataField		(_
						ByRef obj_field, _
						ByVal val_data, _
						ByVal def_data _
						)
	Dim ret_data		: ret_data = val_data
	Dim field_name		: field_name = obj_field.Name
	Dim field_type		: field_type = obj_field.Type


	Select Case field_type
		'------------------------------------------------
		'- Boolean
		'------------------------------------------------
		'[11] adBoolean
		Case 11
			Select Case LCase(Trim(val_data))
				Case "true", "1","y"
					ret_data = TRUE
				Case "false", "0","n"
					ret_data = False
				Case Else
					If chkEmptyGlobal(def_data) = FALSE Then
						ret_data = def_data
					Else
						ret_data = NULL
					End If
			End Select
		'------------------------------------------------
		'- Integer & Decimal
		'------------------------------------------------
		'Integer
		'[2] adSmallInt
		'[3] adInteger
		'[16] adTinyInt
		'[17] adUnsignedTinyInt
		'[18] adUnsignedSmallInt
		'[19] adUnsignedInt
		'[20] adBigInt
		'[21] adUnsignedBigInt

		'Decimal
		'[4] adSingle
		'[5] adDouble
		'[6] adCurrency
		'[14] adDecimal
		'[131] adNumeric
		Case 2, 3, 16, 17, 18, 19, 20, 21, 4, 5, 6, 14, 131
			If IsNumeric(val_data) = TRUE Then
				ret_data = CDbl(val_data)
			Else
				If chkEmptyGlobal(def_data) = FALSE Then
					If IsNumeric(def_data) = TRUE Then
						ret_data = CDbl(def_data)
					Else
						ret_data = NULL
					End If
				Else
					ret_data = NULL
				End If
			End If
		'------------------------------------------------
		'- Date
		'------------------------------------------------
		'[7] adDate
		'[133] adDBDate
		'[135] adDBTimeStamp
		Case  7, 133, 135
			If IsDate(val_data) = TRUE Then
				ret_data = val_data
			Else
				If chkEmptyGlobal(def_data) = FALSE Then
					If IsDate(def_data) = TRUE Then
						ret_data = def_data
					Else
						ret_data = NULL
					End If
				Else
					ret_data = NULL
				End If
			End If
		'------------------------------------------------
		'- String
		'------------------------------------------------
		Case Else
			If chkEmptyGlobal(val_data) = FALSE Then
				ret_data = val_data
			Else
				If chkEmptyGlobal(def_data) = FALSE Then
					ret_data = def_data
				Else
					ret_data = val_data
				End If
			End If
	End Select
	'Response.Write field_name & ":" & ret_data & "<BR>"
	obj_field.Value = ret_data
End Sub
'==================================================================================================
'= Name: SetTableField
'= Desc: 
'==================================================================================================
Sub SetTableField	(_
					ByRef obj_field, _
					ByVal value _
					)
	'On Error Resume Next

	Dim flg_update : flg_update = False
	Dim column
	Dim column_type
	Dim data_before
	Dim data_after
	Dim user_change : user_change = ""

	column		= obj_field.Name
	column_type	= obj_field.Type
	data_before = obj_field.Value
	data_after	= value

	Select Case column_type

		'------------------------------------------------
		'- Boolean
		'------------------------------------------------
		'[11] adBoolean
		Case 11
			If chkEmptyGlobal(data_after) = TRUE Then
				If chkEmptyGlobal(data_before) = FALSE Then
					flg_update = TRUE
					data_after = NULL
				End If
			Else
				If chkEmptyGlobal(data_before) = TRUE Then
					flg_update = TRUE
				Else
					If IsNumeric(data_after) = TRUE Then
						If CBool(data_before) <> CBool(data_after) Then
							flg_update = True
							data_after = CBool(data_after)
						End If
					End If
				End If
			End If
		'------------------------------------------------
		'- Integer
		'------------------------------------------------
		'[2] adSmallInt
		'[3] adInteger
		'[16] adTinyInt
		'[17] adUnsignedTinyInt
		'[18] adUnsignedSmallInt
		'[19] adUnsignedInt
		'[20] adBigInt
		'[21] adUnsignedBigInt
		Case 2, 3, 16, 17, 18, 19, 20, 21
			If chkEmptyGlobal(data_after) = TRUE Then
				If chkEmptyGlobal(data_before) = FALSE Then
					flg_update = TRUE
					data_after = NULL
				End If
			Else
				If chkEmptyGlobal(data_before) = TRUE Then
					flg_update = TRUE
				Else
					If IsNumeric(data_after) = TRUE Then
						If CDbl(data_before) <> CDbl(data_after) Then
							flg_update = TRUE
						End If
					End If
				End If
			End If
		'------------------------------------------------
		'- Decimal
		'------------------------------------------------
		'[4] adSingle
		'[5] adDouble
		'[6] adCurrency
		'[14] adDecimal
		'[131] adNumeric
		Case 4, 5, 6, 14, 131
			
			If chkEmptyGlobal(data_after) = TRUE Then
				If chkEmptyGlobal(data_before) = FALSE Then
					flg_update = TRUE
					data_after = NULL
				End If
			Else
				If chkEmptyGlobal(data_before) = TRUE Then
					flg_update = TRUE
				Else
					If IsNumeric(data_after) = TRUE Then
						If CDbl(data_before) <> CDbl(data_after) Then
							flg_update = TRUE
						End If
					End If
				End If
			End If
		'------------------------------------------------
		'- Date
		'------------------------------------------------
		'[7] adDate
		'[133] adDBDate
		'[135] adDBTimeStamp
		Case  7, 133, 135
			If chkEmptyGlobal(data_after) = TRUE Then
				If chkEmptyGlobal(data_before) = FALSE Then
					flg_update = TRUE
					data_after = NULL
				End If
			Else
				If chkEmptyGlobal(data_before) = TRUE Then
					flg_update = TRUE
				Else
					If IsDate(data_after) = TRUE Then
						If CDate(data_before) <> CDate(data_after) Then
							flg_update = True
							data_after = CDate(data_after)
						End If
					End If
				End If
			End If
		'------------------------------------------------
		'- String
		'------------------------------------------------
		Case Else
			'Response.Write "data_after:" & data_after	& "<BR>"
			If chkEmptyGlobal(data_after) = TRUE Then
				If chkEmptyGlobal(data_before) = FALSE Then
					flg_update = TRUE
				End If
			Else
				If chkEmptyGlobal(data_before) = TRUE Then
					flg_update = TRUE
				Else
					If data_before <> data_after Then
						flg_update = TRUE
					End If
				End If
			End If

			' obj_field.ActualSize
			'Response.Write TypeName(data_after) & ":" & data_after & vbCrLf 
			If flg_update = TRUE Then
				data_after = data_after
			End If
	End Select
	
	If flg_update = TRUE Then
		'If LCase(Session("userid")) = "jcoh" Then
		'Response.Write "column:" & column & "<br>"
		'Response.Write "data_after:" & data_after & "<br>"
		'End If
		obj_field.Value = data_after
	End If
End Sub

'==================================================================================================
'= Name: checkSQLInjection
'= Desc: 
'==================================================================================================
Function checkSQLInjection (ByVal strWords)
	Dim badChars, newChars, tmpChars, regEx, i

	If chkEmptyGlobal(strWords) = FALSE Then
		'-----------------------------------------------
		'-
		'-----------------------------------------------
		badChars = array( _
		"select(.*)(from|with|by){1}", _
		"insert(.*)(into|values){1}", _
		"update(.*)set", "delete(.*)(from|with){1}", _
		"drop(.*)(from|aggre|role|assem|key|cert|cont|credential|data|endpoint|event|f ulltext|function|index|login|type|schema|procedure|que|remote|role|route|sign| stat|syno|table|trigger|user|view|xml){1}", _
		"alter(.*)(application|assem|key|author|cert|credential|data|endpoint|fulltext |function|index|login|type|schema|procedure|que|remote|role|route|serv|table|u ser|view|xml){1}", _
		"xp_", _
		"sp_", _
		"restore\s", _
		"grant\s", _
		"revoke\s", _
		"dbcc", _
		"dump", _
		"use\s", _
		"set\s", _
		"truncate\s", _
		"backup\s", _
		"load\s", _
		"save\s", _
		"shutdown", _
		"cast(.*)\(", _
		"convert(.*)\(", _
		"execute\s", _
		"updatetext", _
		"writetext", _
		"reconfigure", _
		"/\*", _
		"\*/", _
		";", _
		"\-\-", _
		"\[", _
		"\]", _
		"char(.*)\(", _
		"nchar(.*)\("_
		) 
		
		'-----------------------------------------------
		'-
		'-----------------------------------------------
		newChars = strWords

		For i = 0 To uBound(badChars)
			Set regEx = New RegExp
			regEx.Pattern = badChars(i)
			regEx.IgnoreCase = True
			regEx.Global = True
			newChars = regEx.Replace(newChars, "")
			Set regEx = nothing
		Next ' For i = 0 To uBound(badChars)

		newChars = replace(newChars, "'", "''")
	End If
	checkSQLInjection = newChars
End Function



	


'==================================================================================================
'= Name: CnvrPhoneNumber
'= Desc: 
'= Return : 
'==================================================================================================
Function CnvrPhoneNumber (ByVal phone_number)
	Dim ret_val : ret_val = ""

	Dim tmp_str			: tmp_str = ""
	Dim tmp_phone		: tmp_phone_number = ""
	Dim tmp_phone_pre	: tmp_phone_pre = ""
	Dim objRegExp

	tmp_str = Trim(phone_number)

	If tmp_str <> "" Then
		Set objRegExp = New Regexp

		objRegExp.IgnoreCase = True
		objRegExp.Global = True
		objRegExp.Pattern = "((?![0-9]).)+"
		tmp_str = objRegExp.Replace(tmp_str, "")
		Set objRegExp = Nothing
	End If 

	CnvrPhoneNumber = tmp_str
End Function
'==================================================================================================
'= Name: RequestToServer
'= Desc: 
'= Return : 
'==================================================================================================
Function RequestToServer (ByVal url,ByVal data,ByVal response_type)
	Dim objXMLHTTP 
	If chkEmptyGlobal(url) = FALSE AND chkEmptyGlobal(response_type) = FALSE Then
		Set objXMLHTTP = Server.Createobject("MSXML2.ServerXMLHTTP")
		'objXMLHTTP.setTimeouts(2000,2000,2000,2000)
		'objXMLHTTP.setOption 2, 13056 ' https, ssl
		objXMLHTTP.Open "POST",url,false
		objXMLHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXMLHTTP.send data

		'Response.ContentType = "text/xml"

		Select Case LCase(Trim(response_type))
			Case "txt"
				RequestToServer = objXMLHTTP.responseText
			Case "xml"
				Set RequestToServer = objXMLHTTP.responseXML
		End Select
		Set objXMLHTTP = Nothing
	End If
End Function
'/**
'* Returns one of two objects, depending on the evaluation of an expression.
'* @name	IIF
'* @function
'* @param	{boolean} psdStr - A valid Boolean expression
'* @param	{*} trueStr - Value to return if boolean_expression evaluates to true.
'* @param	{*} falseStr - Value to return if boolean_expression evaluates to false.
'*/

'==================================================================================================
'= Name: IIF
'= Desc: Returns one of two objects, depending on the evaluation of an expression.
'==================================================================================================
Public Function IIF(ByVal psdStr, ByVal trueStr, ByVal falseStr)
  If psdStr Then
    IIF = trueStr
  Else
    IIF = falseStr
  End If
End Function

Public Function UCWords (ByVal val)
	UCWords = IIF(chkEmptyGlobal(val),"",UCase(Left(val,1)) & LCase(Right(val, Len(val) - 1)))
End Function
'/**
'* Indicates a specified variable is empty, null or null string.
'* @name	IsBlank
'* @param	{*} var	Any type
'*/
'==================================================================================================
'= Name: IsBlank
'= Desc: 
'==================================================================================================
Public Function IsBlank(ByVal var)
	
	If IsArray(var) Then
		IsBlank = Choice (UBound(var)=-1, TRUE, FALSE)
	Else
		var = Trim(var)
		IsBlank = (IsEmpty(var) OR IsNull(var) OR var = "")
	End If
End Function
'/**
'* Indicates a specified variable has value or not.
'* @name	IsNotBlank
'* @param	{*} var	Any type
'*/
'==================================================================================================
'= Name: IsNotBlank
'= Desc: 
'==================================================================================================
Public Function IsNotBlank(ByVal var)
	IsNotBlank = Not (IsBlank(var))
End Function

'==================================================================================================
'= Name: ExportFileStream
'= Desc: 
'==================================================================================================
Sub ExportFileStream (ByVal file_path, ByVal file_name_extention)
	Dim BUFFERSIZE : BUFFERSIZE = 1024
	Dim fs
	Dim File_Size
	Dim objFile
	Dim adoStream
	Dim tot_size
	Dim read_cnt
	Dim read_size
	Dim i
	set fs = Server.CreateObject("Scripting.FileSystemObject")
	If fs.FileExists(file_path & file_name_extention) Then
		Set objFile = fs.GetFile(file_path & file_name_extention)
		File_Size = objFile.Size
		Set objFile = Nothing
		
		Response.Buffer = False
		Server.ScriptTimeout = 300000
		Response.ContentType = "application/x-unknown"
		Response.AddHeader "Content-Disposition","attachment; filename="& file_name
		
		Set adoStream = CreateObject("ADODB.Stream")
		adoStream.Open()
		adoStream.Type = 1

		adoStream.LoadFromFile(file_path & file_name_extention)
		
		tot_size = adoStream.Size
		read_cnt = tot_size / BUFFERSIZE
		
		if tot_size < BUFFERSIZE then 
			read_size = tot_size
		else
			read_size = BUFFERSIZE
		end if
		
		tot_read = 0
		For i = 0 to read_cnt
			if (tot_read+read_size) > tot_size then
				read_size = tot_size - tot_read
			end if
			Response.BinaryWrite adoStream.Read(read_size) 
		
			tot_read = tot_read + read_size
			if tot_read >= tot_size then exit for
		Next

		adoStream.Close 
		Set adoStream = Nothing 
		
	end if
	set fs = nothing
End Sub

'--------------------------------------------------------------------------------------------------
'- CDictionary: (args: Any)
'--------------------------------------------------------------------------------------------------
Function CDictionary (ByVal args)
	Dim o_dic 
	Select Case TypeName(args)
		Case "Variant()"
			Dim i
			Set o_dic = CreateObject("Scripting.Dictionary")
			If ((UBound(args) + 1) mod 2) = 0 Then
				
				For i = 0 to UBound(args) step 2
					o_dic.Add args(i), args(i+1)
				Next
			End If 

			Set CDictionary = o_dic
			Set o_dic = Nothing
		Case "Recordset"
			Dim field
			Set o_dic = CreateObject("Scripting.Dictionary")

			If Not args.EOF Then
				For Each field In args.Fields
					o_dic.Add field.Name, field.Value
				Next
			End If
			
			Set CDictionary = o_dic
			Set o_dic = Nothing
		Case "Dictionary"
			Set CDictionary = args
		Case Else
			Set CDictionary = CreateObject("Scripting.Dictionary")
	End Select
End Function

%>