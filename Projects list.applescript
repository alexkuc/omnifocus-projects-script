tell application "OmniFocus"
	set currentDate to current date
	set allProjects to a reference to flattened projects of default document
	set listProjects to allProjects whose status is active and (defer date > currentDate or defer date is missing value)
end tell

tell application "OmniOutliner"
	activate
	set template to ((path to application support from user domain as string) & "The Omni Group:OmniOutliner:Templates:" & "Blank.oo3template:")
	open template
	tell front document
		repeat with thisProject in listProjects
			set urlProject to "omnifocus:///task/" & (id of thisProject)
			set nameProject to name of thisProject
			set newRow to make new row with properties {topic:nameProject} at end of (parent of last row)
			set value of attribute named "link" of style of newRow to urlProject
		end repeat
		delete first row
	end tell
end tell
