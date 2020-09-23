<div align="center">

## ADODBTypeName


</div>

### Description

Quickly find out the type of variables returned from your adodb recordset. TypeName() function doesn't do the trick. Databases offer additional data types. This script helps solve type problems without having to lookup the name of the numbers returned.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Lewis E\. Moten III](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/lewis-e-moten-iii.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Data Structures](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/data-structures__4-8.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/lewis-e-moten-iii-adodbtypename__4-7470/archive/master.zip)





### Source Code

```
Response.write ADODBTypeName(Rs.Fields("UserID").Type)
Function ADODBTypeName(ByRef plngType)
	Select Case plngType
		Case 0:ADODBTypeName="adEmpty"
		Case 2:ADODBTypeName="adSmallInt"
		Case 3:ADODBTypeName="adInteger"
		Case 4:ADODBTypeName="adSingle"
		Case 5:ADODBTypeName="adDouble"
		Case 6:ADODBTypeName="adCurrency"
		Case 8:ADODBTypeName="adBSTR"
		Case 9:ADODBTypeName="adDispatch"
		Case 10:ADODBTypeName="adError"
		Case 11:ADODBTypeName="adBoolean"
		Case 12:ADODBTypeName="adVariant"
		Case 13:ADODBTypeName="adIUnknown"
		Case 14:ADODBTypeName="adDecimal"
		Case 16:ADODBTypeName="adTinyInt"
		Case 17:ADODBTypeName="adUnsignedTinyInt"
		Case 18:ADODBTypeName="adUnsignedSmallInt"
		Case 19:ADODBTypeName="adUnsignedInt"
		Case 21:ADODBTypeName="adUnsignedBigInt"
		Case 64:ADODBTypeName="adFileTime"
		Case 72:ADODBTypeName="adGUID"
		Case 20:ADODBTypeName="adBigInt"
		Case 128:ADODBTypeName="adBinary"
		Case 129:ADODBTypeName="adChar"
		Case 130:ADODBTypeName="adWChar"
		Case 131:ADODBTypeName="adNumeric"
		Case 132:ADODBTypeName="adUserDefined"
		Case 133:ADODBTypeName="adDBDate"
		Case 134:ADODBTypeName="adDBTime"
		Case 135:ADODBTypeName="adDBTimeStamp"
		Case 136:ADODBTypeName="adChapter"
		Case 137:ADODBTypeName="adDBFileTime"
		Case 138:ADODBTypeName="adPropVariant"
		Case 139:ADODBTypeName="adVarNumeric"
		Case 200:ADODBTypeName="adVarChar"
		Case 201:ADODBTypeName="adLongVarChar"
		Case 202:ADODBTypeName="adVarWChar"
		Case 203:ADODBTypeName="adLongVarWChar"
		Case 204:ADODBTypeName="adVarBinary"
		Case 205:ADODBTypeName="adLongVarBinary"
	End Select
End Function
```

