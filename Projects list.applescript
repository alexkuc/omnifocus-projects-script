use application "OmniFocus"

use scripting additions

tell application "OmniOutliner"
	
	if version > 5 then
		
		set template to ((path to application support from user domain as string) & "The Omni Group:OmniOutliner:Pro Templates:" & "Blank.otemplate:") --- template path for OO5
		
	else
		
		set template to ((path to application support from user domain as string) & "The Omni Group:OmniOutliner:Templates:" & "Blank.oo3template:") --- template for OO3 or OO4
		
	end if
	
	
	set currentDate to current date
	
	activate
	
	open template
	
	tell front document
		
		set title of second column to "Project Name"
		
		make new column with properties {title:"Project Status", sort order:ascending}
		
		make new column with properties {title:"Project Due Date", column type:datetime}
		
		repeat with thisProject in (flattened projects of default document whose status is active)
			
			using terms from application "OmniFocus" --- workaround for "folder" term collision (future use)
				
				
				
			end using terms from
			
			if text of second cell of first row is equal to "" then
				
				set newRow to first row
				
			else
				
				set newRow to make new row at end of (parent of last row)
				
			end if
			
			set text of second cell of newRow to name of thisProject
			
			if (defer date of thisProject) is greater than currentDate then
				
				set statusProject to "deferred"
				
			else if ((number of tasks of thisProject) - (number of completed tasks of thisProject)) is equal to 0 then
				
				set statusProject to "stalled"
				
			else
				set statusProject to status of thisProject as string
				
			end if
			
			set value of attribute named "link" of style of text of second cell of newRow to "omnifocus:///task/" & id of thisProject
			
			set text of third cell of newRow to statusProject
			
			if due date of thisProject is not equal to missing value then
				
				set value of fourth cell of newRow to due date of thisProject as date
				
			end if
			
		end repeat
		
	end tell
	
end tell
