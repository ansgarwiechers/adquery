What is this?
=============
I got tired of writing the same boilerplate code over and over again whenever I
needed to query some Active Directory for user, computer or group information.
This script wraps the initialization of ADODB.Connection and ADODB.Command in
a class, sets some (IMHO reasonable) default values and provides properties so
that the usual parameters can be adjusted as the situation requires.

The class does not do any error handling by itself. All errors (particularly
errors during class initialization) MUST be handled by the parent script.


Copyright
=========
This script is distributed according to the terms of the GNU General Public
License Version 2.0 as found at
<http://www.gnu.org/licenses/old-licenses/gpl-2.0.html>.

This program is distributed in the hope that it will be useful, but WITHOUT ANY
WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A
PARTICULAR PURPOSE.  See the GNU General Public License for more details.


Including the class
===================
To use the class in your scripts you must either copy/paste it to the script,
or use this neat import procedure:

' <http://gazeek.com/coding/importing-vbs-files-in-your-vbscript-project/>
Sub Import(ByVal filename)
	Dim fso, sh, file, code

	Set fso = CreateObject("Scripting.FileSystemObject")
	Set sh = CreateObject("WScript.Shell")
	filename = sh.ExpandEnvironmentStrings(filename)
	filename = fso.GetAbsolutePathName(filename)
	Set file = fso.OpenTextFile(filename)
	code = file.ReadAll
	file.Close
	ExecuteGlobal(code)
End Sub


Examples
========
' create a new ADQuery instance
Set qry = New ADQuery

' create a new ADQuery instance and initialize it with custom values
base   = "ou=servers,dc=example,dc=com"
filter = "(objectCategory=computer)"
attr   = Array("sAMAccountName", "operatingSystem")
scope  = 1
Set qry = (New ADQuery)(base, filter, attr, scope)

' change query properties
qry.SearchBase   = "ou=foo,dc=sub,dc=example,dc=org"
qry.Filter       = "(&(sAMAccountName>=f)(sAMAccountName<h))"
qry.Attributes   = Array("sn", "givenName")
qry.Scope        = 1
qry.CacheResults = True

' run the query and process the results
Set rs = qry.Execute
If Not rs Is Nothing Then
  Do Until rs.EOF
    WScript.Echo rs.Fields("sn").Value
    rs.MoveNext
  Loop
End If
rs.Close
Set rs = Nothing
