Option Explicit

Dim App				' The Syllabus+ Application.
Dim Coll			' The root Institution object.
Dim myTemplate		' The Activity Template object we are currently working with.
Dim myActivity		' The Activity object we are currently working with.
Dim myStudentSet	' The Student Set object we are currently working with.
Dim templateGroup	' Contains the Activity Template Group we will be allocating.
Dim myResources		' Resource object used for attaching Student Sets to an Activity.
Dim myModule		' Module object a given Activity Template belongs to.

' Connect to the Syllabus+ application and get a reference to the root institution.
' Note that we are assuming the Prog ID is "Splus".
Set App = CreateObject("Splus.application")
Set Coll = App.ActiveCollege

' Create an Activity Template Group to contain the Activity Templates we will be allocating.
Set templateGroup = Coll.CreateActivityTemplateGroup
templateGroup.Name = "1:1 Allocations - " + FormatDateTime(Now()) 

' Now iterate through all the Activity Templates in Syllabus+.
For Each myTemplate in Coll.ActivityTemplates

	' Check that the Activity Template has just 1 Activity.
	If myTemplate.LinkedActivities.Count = 1 Then
	
		' Check that the Activity Template has an associated Module.
		If Not myTemplate.Module Is Nothing Then
		
			Set myActivity = myTemplate.LinkedActivities.Item(1)
			' Are there unallocated Student Sets on this Activity Template?
			If myActivity.RealSize < myTemplate.Module.RealSize Then
				Call templateGroup.Members.Add(myTemplate)
			End If
			
		End If
	
	End If
	
Next	' End of loop through Activity Templates.

' We now have an Activity Template Group containing Activity Templates who only have a single
' Activity object, and the real size of that Activity is less than the Real Size of the 
' Module the Activity and Activity Template belong to.

' Loop over the Activity Template objects in our group.
For Each myTemplate in templateGroup.Members

	Set myActivity = myTemplate.LinkedActivities.Item(1)
	Set myModule = myTemplate.Module
	Set myResources = Coll.CreateResources
	
	For Each myStudentSet in myModule.StudentSets
		Call myResources.Add(myStudentSet)
	Next	' End of loop through Student Sets on the Module.
	
	' Note that the first parameter of the next call is the type of Resource being added; 3 = Student Set.
	' Note that the second parameter of the next call is how the Resources should be added; 1 = Preset.
	Call myActivity.SetResourceRequirement(3, 1, myResources.Count, myResources)

Next	' End of loop through Activity Templates.

' All done, now write back the changes.
Call App.SDB.WriteBack
