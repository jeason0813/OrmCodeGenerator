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


Friend Class cPostazione

    Private _pathApplicazione As String = System.AppDomain.CurrentDomain.BaseDirectory
    Public ReadOnly Property PathApplicazione As String
        Get
            Return _pathApplicazione
        End Get
    End Property

    Friend ReadOnly Property Version() As String
        Get

            Return My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Build & "." & My.Application.Info.Version.Revision

        End Get
    End Property

    Friend Sub ScriviLogFile(ByVal NomeFile As String, ByVal Testo As String)

        Dim objWriter As New System.IO.StreamWriter(NomeFile, True)

        objWriter.WriteLine(Testo)

        objWriter.Close()

    End Sub

End Class
