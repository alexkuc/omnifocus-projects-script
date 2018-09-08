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
		
		make new column with properties {title:"Root Folder"}
		
		repeat with thisProject in (flattened projects of default document whose (status is not done) and (status is not dropped))
			
			using terms from application "OmniFocus" --- workaround for "folder" term collision (future use)
				
				set rootFolderName to folder of thisProject
				
				if rootFolderName is missing value then --- checking for root level projects
					
					set rootFolderName to "n/a"
					
					set foundRootFolderName to true
					
				else
					
					set foundRootFolderName to false
					
				end if
				
				repeat until foundRootFolderName is true
					
					set upperFolder to container of rootFolderName
					
					if name of upperFolder is equal to "OmniFocus" then
						
						set rootFolderName to name of rootFolderName as string
						
						set foundRootFolderName to true
						
					else
						
						set rootFolderName to upperFolder
						
					end if
					
				end repeat
				
			end using terms from
			
			if text of second cell of first row is equal to "" then
				
				set newRow to first row
				
			else
				
				set newRow to make new row at end of (parent of last row)
				
			end if
			
			set text of second cell of newRow to name of thisProject
			
			set statusProject to status of thisProject as string
			
			if statusProject is not equal to "on hold" then
				
				if ((number of tasks of thisProject) - (number of completed tasks of thisProject)) is equal to 0 then
					
					set statusProject to "stalled"
					
				else if (defer date of thisProject) is greater than currentDate then
					
					set statusProject to "deferred"
					
				else if number of available tasks of thisProject is less than 1 then
					
					set statusProject to "deferred"
					
				else if (number of available tasks of thisProject is greater than or equal to 1) and (singleton action holder of thisProject as boolean is false) then
					
					set i to 0
					
					if (sequential of thisProject as boolean is true) then
						
						if (name of context of next task of thisProject as string) contains "wait" then
							
							set i to i + 1
							
						end if
						
					else
						
						repeat with thisTask in (every task of thisProject whose completed is not true)
							
							if (name of context of thisTask as string) contains "wait" then
								
								set i to i + 1
								
							end if
							
						end repeat
						
					end if
					
					if i is equal to number of available tasks of thisProject then
						
						set statusProject to "deferred"
						
					end if
					
				end if
				
			end if
			
			set value of attribute named "link" of style of text of second cell of newRow to "omnifocus:///task/" & id of thisProject
			
			set text of third cell of newRow to statusProject
			
			if due date of thisProject is not equal to missing value then
				
				set value of fourth cell of newRow to due date of thisProject as date
				
			end if
			
			set text of fifth cell of newRow to rootFolderName as string
			
		end repeat
		
	end tell
	
end tell
