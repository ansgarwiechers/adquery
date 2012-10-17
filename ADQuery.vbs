' This program is free software; you can redistribute it and/or
' modify it under the terms of the GNU General Public License
' as published by the Free Software Foundation; either version 2
' of the License, or (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program; if not, write to the Free Software
' Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.

'! Query objects from Active Directory via LDAP.
'!
'! @author  Ansgar Wiechers <ansgar.wiechers@planetcobalt.net>
'! @date    2012-10-17
'! @version 1.0
Class ADQuery
	Private classname_
	Private ADS_SCOPEENUM

	Private conn_
	Private cmd_

	Private base_
	Private filter_
	Private attributes_
	Private scope_

	' == c'tor/d'tor ==========================================================

	'! @brief Constructor.
	'!
	'! Open an ADODB connection to the default naming context and initialize the
	'! ADQuery object with default values.
	'!
	'! @see http://msdn.microsoft.com/en-us/library/windows/desktop/ms681519%28v=vs.85%29.aspx (ADO Connection Object)
	'! @see http://msdn.microsoft.com/en-us/library/windows/desktop/ms677502%28v=vs.85%29.aspx (ADO Command Object)
	'!
	'! @raise Domain not available (0x8007054b)
	Private Sub Class_Initialize()
		Dim errno  : errno  = 0
		Dim errtxt : errtxt = Empty

		Set conn_ = Nothing
		Set cmd_ = Nothing

		classname_ = "ADQuery"
		ADS_SCOPEENUM = Array("base", "onelevel", "subtree")

		On Error Resume Next
		Dim rootDSE : Set rootDSE = GetObject("LDAP://RootDSE")
		If Err Then
			errno  = Err.Number
			errtxt = Err.Description
			If errno = &h8007054b Then errtxt = "Domain not available."
		End If
		On Error Goto 0
		If errno <> 0 Then
			Err.Raise errno, classname_, errtxt
			WScript.Quit 1
		End If

		base_       = "<LDAP://" & rootDSE.Get("defaultNamingContext") & ">"
		filter_     = "(&(objectClass=user)(objectCategory=Person))"
		attributes_ = "distinguishedName"
		scope_      = 2

		Set rootDSE = Nothing

		Set conn_ = CreateObject("ADODB.Connection")
		conn_.Provider = "ADsDSOObject"
		conn_.Open "Active Directory Provider"

		Set cmd_ = CreateObject("ADODB.Command")
		Set cmd_.ActiveConnection = conn_
		cmd_.Properties("Page Size")     = 1000
		cmd_.Properties("Timeout")       = 30
		cmd_.Properties("Cache Results") = False
	End Sub

	'! @brief Destructor.
	'!
	'! Close the connection and clean up when the object is destroyed.
	Private Sub Class_Terminate()
		If Not conn_ Is Nothing Then
			conn_.Close
			Set conn_ = Nothing
		End If
		If Not cmd_ Is Nothing Then Set cmd_ = Nothing
	End Sub

	'! Initialization method to allow for setting all main properties at the
	'! same time. Null values leave the respective property unchanged.
	'!
	'! @param  base    The LDAP search base.
	'! @param  filter  LDAP filter to restrict the returned objects.
	'! @param  attr    The attributes to be retrieved by the query.
	'! @param  scope   The scope for the query.
	'! @return The initialized object.
	Public Default Function Init(base, filter, attr, scope)
		If Not (IsNull(base) Or VarType(base) = vbError) Then Me.SearchBase = base
		If Not (IsNull(filter) Or VarType(filter) = vbError) Then Me.Filter = filter
		If Not (IsNull(attr) Or VarType(attr) = vbError) Then Me.Attributes = attr
		If Not (IsNull(scope) Or VarType(scope) = vbError) Then Me.Scope = scope
		Set Init = Me
	End Function

	' == properties ===========================================================

	'! The search base for the query. Must be an LDAP distinguished name (e.g.
	'! "ou=foo,dc=example,dc=com"). Default is the default naming context of the
	'! computer's domain.
	'!
	'! @raise Malformed distinguished name (450)
	Public Property Get SearchBase()
		SearchBase = Mid(base_, 9, Len(base_)-9)
	End Property

	Public Property Let SearchBase(arg)
		If Not HasValidSyntax(arg) Then
			Err.Raise 450, classname_, "Malformed distinguished name: " & arg
			Exit Property
		End If
		base_ = "<LDAP://" & arg & ">"
	End Property

	'! LDAP filter for the query. The default is to query for user objects.
	'!
	'! @see http://msdn.microsoft.com/en-us/library/windows/desktop/aa746475%28v=vs.85%29.aspx (Search Filter Syntax)
	'!
	'! @raise Argument must be a string (450)
	Public Property Get Filter
		Filter = filter_
	End Property

	Public Property Let Filter(arg)
		If Not IsValidFilter(arg) Then
			Err.Raise 450, classname_, "Argument must be a string."
			Exit Property
		End If
		filter_ = arg
	End Property

	'! Array with the names of the attributes the query should return. The
	'! default is "distinguishedName".
	'!
	'! @raise Invalid argument (450)
	Public Property Get Attributes
		Attributes = Split(attributes_, ",")
	End Property

	Public Property Let Attributes(arg)
		If Not IsAttributeList(arg) Then
			Err.Raise 450, classname_, "Invalid argument. Must be an array of attribute names."
			Exit Property
		End If
		attributes_ = Join(arg, ",")
	End Property

	'! The scope of the query. Valid scopes are defined in the ADS_SCOPEENUM
	'! enumeration:
	'!
	'! - 0 = base
	'! - 1 = onelevel
	'! - 2 = subtree
	'!
	'! The default is 2 (subtree).
	'!
	'! @see http://msdn.microsoft.com/en-us/library/windows/desktop/aa772286%28v=vs.85%29.aspx (ADS_SCOPEENUM enumeration)
	'!
	'! @raise Invalid scope (450)
	Public Property Get Scope
		Scope = scope_
	End Property

	Public Property Let Scope(arg)
		If Not IsValidScope(arg) Then
			Err.Raise 450, classname_, "Invalid scope."
			Exit Property
		End If
		scope_ = arg
	End Property

	'! Number of records per logical page of data. Must be a postivie integer.
	'! The default is 1000.
	'!
	'! @see http://msdn.microsoft.com/en-us/library/windows/desktop/ms681431%28v=vs.85%29.aspx (PageSize Property)
	'!
	'! @raise Invalid argument (450)
	Public Property Get PageSize
		PageSize = cmd_.Properties("Page Size")
	End Property

	Public Property Let PageSize(arg)
		If Not IsPositiveInteger(arg) Then
			Err.Raise 450, classname_, "Invalid argument. Must be a positive integer."
			Exit Property
		End If
		cmd_.Properties("Page Size") = arg
	End Property

	'! Timeout for the AD command in seconds. Must be a positive integer. The
	'! default is 30 seconds.
	'!
	'! @see http://msdn.microsoft.com/en-us/library/windows/desktop/ms678265%28v=vs.85%29.aspx (CommandTimeout Property)
	'!
	'! @raise Invalid argument (450)
	Public Property Get Timeout
		PageSize = cmd_.Properties("Timeout")
	End Property

	Public Property Let Timeout(arg)
		If Not IsPositiveInteger(arg) Then
			Err.Raise 450, classname_, "Invalid argument. Must be a positive integer."
			Exit Property
		End If
		cmd_.Properties("Timeout") = arg
	End Property

	'! Indicates whether or not the results of the query should be cached. The
	'! default is False.
	'!
	'! @raise Invalid argument (450)
	Public Property Get CacheResults
		PageSize = cmd_.Properties("Cache Results")
	End Property

	Public Property Let CacheResults(arg)
		If VarType(arg) <> vbBoolean Then
			Err.Raise 450, classname_, "Invalid argument. Must be of type boolean."
			Exit Property
		End If
		cmd_.Properties("Cache Results") = arg
	End Property

	' == validation functions =================================================

	'! Check if the given string has valid LDAP domain name syntax.
	'!
	'! @param  str  LDAP domain name string
	'! @return True if the given string has valid LDAP domain name syntax,
	'!         otherwise False.
	Private Function HasValidSyntax(str)
		Dim re : Set re = New RegExp
		re.Pattern = "^(cn=\S+?\s*,\s*)?(ou=\S+?\s*,\s*)*dc=\S+?(\s*,\s*dc=\S+?)+$"
		re.IgnoreCase = True
		HasValidSyntax = re.Test(str)
		Set re = Nothing
	End Function

	'! Check if the given argument is a valid filter. Right now this only
	'! checks if the argument is a string.
	'!
	'! @param  str  An LDAP filter string.
	'! @return True if the argument is a string, otherwise False.
	Private Function IsValidFilter(str)
		IsValidFilter = False
		If VarType(arg) = vbString Then
			IsValidFilter = True
		End If
	End Function

	'! Check if the given argument is a valid attribute list. To be valid, the
	'! argument must be an array of non-empty string values.
	'!
	'! @param  list   Array with attribute names.
	'! @return True if the given argument is an array of non-empty string
	'!         values, otherwise False.
	Private Function IsAttributeList(list)
		Dim i

		IsAttributeList = False
		If (VarType(list) And vbArray) = vbArray Then
			If UBound(list) >= 0 Then
				IsAttributeList = True
				For i = 0 To UBound(list)
					If VarType(list(i)) <> vbString Then
						IsAttributeList = False
						Exit For
					ElseIf Len(Trim(list(i))) = 0 Then
						IsAttributeList = False
						Exit For
					End If
				Next
			End If
		End If
	End Function

	'! Check if the given argument is a valid scope option as defined per the
	'! ADS_SCOPEENUM enumeration.
	'!
	'! @param  val  A scope option.
	'! @return True if the given argument is valid scope option, otherwise
	'!         False.
	Private Function IsValidScope(val)
		IsValidScope = False
		If VarType(val) = vbInteger Then
			If val >= 0 And val <= 2 Then IsValidScope = True
		End If
	End Function

	'! Check if the given argument is a positive integer.
	'!
	'! @param  val  An integer.
	'! @return True if the given argument is a positive integer, otherwise
	'!         False.
	Private Function IsPositiveInteger(val)
		IsPositiveInteger = False
		If VarType(val) = vbInteger Then
			If val >= 0 Then IsPositiveInteger = True
		End If
	End Function

	' == public methods =======================================================

	'! Run an AD query with the configured parameters.
	'!
	'! @return An ADO recordset with the results returned by the query or Nothing.
	'!
	'! @see http://msdn.microsoft.com/en-us/library/windows/desktop/ms681510%28v=vs.85%29.aspx (ADO Recordset Object)
	Function Execute()
		Dim rs

		' Raising an error in Class_Initialize() doesn't seem to terminate the
		' calling script, so I'm adding this extra check to abort queries when
		' the object wasn't properly initialized.
		If cmd_ Is Nothing Then Err.Raise 424, WScript.ScriptName, "Object required."

		cmd_.CommandText = base_ & ";" & filter_ & ";" & attributes_ & ";" & ADS_SCOPEENUM(scope_)
		Set rs = cmd_.Execute
		If rs.BOF And rs.EOF Then
			' record set is empty
			Set Execute = Nothing
		Else
			rs.MoveFirst
			Set Execute = rs
		End If
	End Function
End Class
