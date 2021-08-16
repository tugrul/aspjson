<%
Function QueryToJSON_(rs, propertyCase)
  Dim jsa, headers, headersBound, rows, rowsBound

  Set jsa = jsArray()
  if (NOT rs.EOF) AND (NOT rs.BOF) then
  	headersBound = rs.fields.Count - 1
  	redim headers(headersBound)

    if (propertyCase = "upper") THEN
      for i = 0 to headersBound
        headers(i) = UCase(rs.fields(i).name)
      Next
    elseif (propertyCase = "lower") THEN
      for i = 0 to headersBound
        headers(i) = LCase(rs.fields(i).name)
      Next
    else
      for i = 0 to headersBound
        headers(i) = rs.fields(i).name
      Next
    end if

    rows = rs.getrows()
    rowsBound = UBound(rows, 2)

    for rowIndex = 0 to rowsBound
      set jsa(null) = jsObject()
      for headerIndex = 0 to headersBound
  			jsa(null)(headers(headerIndex)) = rows(headerIndex, rowIndex)
  		next
    next
  end if

  Set QueryToJSON_ = jsa
End Function

Function QueryToJSON(dbc, sql)
  dim rs
  Set rs = dbc.Execute(sql)
  Set QueryToJSON = QueryToJSON_(rs, "normal")
End Function

Function UCaseQueryToJSON(dbc, sql)
  dim rs
  Set rs = dbc.Execute(sql)
  Set UCaseQueryToJSON = QueryToJSON_(rs, "upper")
End Function

Function LCaseQueryToJSON(dbc, sql)
  dim rs
  Set rs = dbc.Execute(sql)
  Set LCaseQueryToJSON = QueryToJSON_(rs, "lower")
End Function

Function PreparedQueryToJSON(cmd)
  dim rs
  Set rs = cmd.Execute()
  Set PreparedQueryToJSON = QueryToJSON_(rs, "normal")
End Function

Function UCasePreparedQueryToJSON(cmd)
  dim rs
  Set rs = cmd.Execute()
  Set UCasePreparedQueryToJSON = QueryToJSON_(rs, "upper")
End Function

Function LCasePreparedQueryToJSON(cmd)
  dim rs
  Set rs = cmd.Execute()
  Set LCasePreparedQueryToJSON = QueryToJSON_(rs, "lower")
End Function
%>
