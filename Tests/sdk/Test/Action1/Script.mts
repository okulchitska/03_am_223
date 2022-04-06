Option Explicit

Dim MyMsgBox
Set MyMsgBox = DotNetFactory.CreateInstance("System.Windows.Forms.MessageBox", "System.Windows.Forms")

Dim apiUser, apiSecret, octaneUrl
Dim sharedSpaceId, workspaceId, runId

apiUser = Parameter("aApiUser")
apiSecret = Parameter("aApiSecret")
octaneUrl = Parameter("aOctaneUrl")
sharedSpaceId = Parameter("aOctaneSpaceId")
workspaceId = Parameter("aOctaneWorkspaceId")
'runId = Parameter("aRunId")

Dim restConnector, connectionInfo, isConnected
Set restConnector = DotNetFactory.CreateInstance("MicroFocus.Adm.Octane.Api.Core.Connector.RestConnector", "MicroFocus.Adm.Octane.Api.Core")
Set connectionInfo = DotNetFactory.CreateInstance("MicroFocus.Adm.Octane.Api.Core.Connector.UserPassConnectionInfo", "MicroFocus.Adm.Octane.Api.Core", apiUser, apiSecret)
isConnected = restConnector.Connect(octaneUrl, connectionInfo)
'MyMsgBox.Show  isConnected, "Is Connected"

Dim context, entityService
Set context = DotNetFactory.CreateInstance("MicroFocus.Adm.Octane.Api.Core.Services.RequestContext.WorkspaceContext", "MicroFocus.Adm.Octane.Api.Core", sharedSpaceId, workspaceId)
Set entityService = DotNetFactory.CreateInstance("MicroFocus.Adm.Octane.Api.Core.Services.NonGenericsEntityService", "MicroFocus.Adm.Octane.Api.Core", restConnector)


Dim entType, entId, entFields, entFieldsAttach
entType = "test"
entId = "2203"
entFields = Array("id", "subtype", "name", "phase", "automation_status", "source_id_udf", "author")
entFieldsAttach = Array("id", "name")


Dim attachmentsList, attachmentsList1, attachmentsList2, attachmentsName, orderBy, limit, offset
orderBy = "id"
limit = CInt(2)
offset = CInt(0)
Set attachmentsList = entityService.Get(context, "attachments", "(owner_test={id=" + entId + "})", entFieldsAttach)
'Set attachmentsList1 = entityService.Get(context, "attachments", "(owner_test={id=" + entId + "})", entFields, 1, 1)
'Set attachmentsList2 = entityService.Get(context, "attachments", "(owner_test={id=" + entId + "})", entFields, "id", 1, 0)

attachmentsName = ""
Dim i, element
For i = 0 To attachmentsList.BaseEntities.Count - 1
	Set element = attachmentsList.BaseEntities.Item(CInt(i))
	If (Len(attachmentsName) > 0) Then
		attachmentsName = attachmentsName + ", "
	End If
	attachmentsName = attachmentsName + element.Name
	entityService.DownloadAttachment "/api/shared_spaces/" +sharedSpaceId+ "/workspaces/" +workspaceId+ "/attachments/" +element.Id+ "/" + element.Name, "C:\\Downloads\\" +element.Name
Next
'MyMsgBox.Show "Attachments: " + attachmentsName, "Attachments"


'Write to file
Dim test, FSO, outfile
Set test = entityService.GetById(context, entType, entId, entFields)
Set FSO = CreateObject("Scripting.FileSystemObject")
Set outFile = FSO.CreateTextFile("C:\Downloads\MyFile_TestEntity.txt",True)
'outFile.Write "ID: " + test.Id + ", Test Type: " + test.Subtype
outFile.WriteLine "Test Type: " + test.Subtype
outFile.WriteLine vbCrLf & "Octane ID: " + test.Id
outFile.WriteLine "Test Name: " + test.GetValue("name")
outFile.WriteLine "ALM QC ID: " + test.GetValue ("source_id_udf")
outFile.WriteLine "Phase: " + test.Phase.Id
outFile.WriteLine "Automated Status: " + test.GetValue("automation_status").Id
outFile.WriteLine "Author: " + test.GetValue ("author").Name
outFile.WriteLine vbCrLf & "Attachments: " + attachmentsName
outFile.Close
