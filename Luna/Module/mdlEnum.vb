'
' Copyright Diego Lunadei (d.lunadei--at--gmail.com)
'
' Licensed under the Apache License, Version 2.0 (the "License");
' you may not use this file except in compliance with the License.
' You may obtain a copy of the License at
'
'     http://www.apache.org/licenses/LICENSE-2.0
'
' Unless required by applicable law or agreed to in writing, software
' distributed under the License is distributed on an "AS IS" BASIS,
' WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
' See the License for the specific language governing permissions and
' limitations under the License.
'


Module mdlEnum

    Friend Enum enDatasourceType As Integer
        Access
        SQLServer
        SQLServerCompact
        Oracle
        LunaDataSchema
    End Enum

    Friend Enum enLanguage As Integer
        VbNet = 0
        CSharp
    End Enum

    Friend Enum enTableType As Integer
        Table = 0
        View
    End Enum

End Module