use application "OmniFocus"

use scripting additions

tell application "OmniOutliner"
	
	if version ³ 5 then
		
		set template to ((path to application support from user domain as string) & "The Omni Group:OmniOutliner:Pro Templates:" & "Blank.otemplate:") --- template path for OO5
		
	else
		
		set template to ((path to application support from user domain as string) & "The Omni Group:OmniOutliner:Templates:" & "Blank.oo3template:") --- template for OO3 or OO4
		
	end if
	
	set currentDate to current date
	
	activate
	
	open template
	
	tell front document
		
		tell application "OmniFocus"
			
			if version < 3 then
				
				error "You need to have OmniFocus version 3 or later. Early versions are not supported!"
				
			end if
			
		end tell
		
		
		set title of second column to "Project Name"
		
		(*
		
		create columns in OmniOutliner
		
		*)
		
		make new column with properties {title:"Project Status"}
		
		make new column with properties {title:"Project Due Date", column type:datetime}
		
		make new column with properties {title:"Root Folder", sort order:ascending}
		
		(*
		
		iterate through projects which are (must match ALL conditions):
		
		- active
		
		- on hold
		
		- not dropped
		
		- not (effectively) dropped
		
		*)
		
		repeat with thisProject in (flattened projects of default document whose (effective status is active status or effective status is on hold status))
			
			using terms from application "OmniFocus" --- workaround for "folder" term collision (future use)	
				
				(*
				
				find root folder of the project (start section)
				
				*)
				
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
			
			(*
			
			create new row in OmniOutliner
			
			*)
			
			if text of second cell of first row is equal to "" then
				
				set newRow to first row
				
			else
				
				set newRow to make new row at end of (parent of last row)
				
			end if
			
			(*
			
			add OmniFocus project name and create deep link to it
			
			*)
			
			set text of second cell of newRow to name of thisProject
			
			set value of attribute named "link" of style of text of second cell of newRow to "omnifocus:///task/" & id of thisProject
			
			(*
			
			add OmniFocus project status
			
			*)
			
			set projectStatus to status of thisProject as string
			
			--- singleton projects are not given granularity as they simply hold a list of tasks (actions)
			
			if singleton action holder of thisProject as boolean is true then
				
				set projectStatus to "action list"
				
			else
				
				--- regular (active) projects (sequential or parallel) are given high level granurality
				
				if thisProject's status is active status then
					
					--- internal variables used to determine project's status
					
					set cAvailTasks to 0
					
					set cAvailTasksTags to 0
					
					set cRemainTasks to 0
					
					set cRemainTasksTags to 0
					
					set cWaitingTasks to 0
					
					set cDeferredTasks to 0
					
					set bFirstAvailTaskIsWaiting to false
					
					--- iterate through every remaining task of the project
					
					repeat with thisTask in (flattened tasks of thisProject whose effectively completed is false and effectively dropped is false)
						
						set cRemainTasks to cRemainTasks + 1
						
						if thisTask's primary tag is not missing value then
							
							set cRemainTasksTags to cRemainTasksTags + 1
							
							if thisTask's primary tag's name contains "wait" then
								
								set cWaitingTasks to cWaitingTasks + 1
								
							end if
							
						end if
						
						if thisTask's defer date ² currentDate and thisTask's blocked is false then
							
							set cAvailTasks to cAvailTasks + 1
							
							if thisTask's primary tag is not missing value then
								
								set cAvailTasksTags to cAvailTasksTags + 1
								
							end if
							
						else
							
							set cDeferredTasks to cDeferredTasks + 1
							
						end if
						
					end repeat
					
					--- check if first available task is "wait-for" for sequential projects
					
					set tmpVar to missing value
					
					if thisProject is sequential and next task of thisProject is not missing value then
						
						if primary tag of next task of thisProject is not missing value then
							
							if name of primary tag of next task of thisProject is not missing value then
								
								set tmpVar to name of primary tag of next task of thisProject
								
							end if
							
						end if
						
					end if
					
					if tmpVar contains "wait" then
						
						set bFirstAvailTaskIsWaiting to true
						
					end if
					
					--- project has no remaining tasks
					
					if cRemainTasks = 0 then
						
						set projectStatus to "stalled (no tasks)"
						
					end if
					
					--- project has remaining tasks with no tags
					
					if cRemainTasks > 0 and cRemainTasksTags = 0 then
						
						set projectStatus to "stalled (no tags)"
						
					end if
					
					--- all project's available tasks are "wait-for" context
					
					if cWaitingTasks > 0 and (cAvailTasksTags = cWaitingTasks or bFirstAvailTaskIsWaiting is true) then
						
						set projectStatus to "waiting"
						
					end if
					
					--- project has no available tasks but has remaining tasks
					
					if thisProject's defer date is missing value and cAvailTasksTags = 0 and cRemainTasksTags > 0 then
						
						set projectStatus to "deferred (tasks)"
						
					end if
					
					--- project is deferred by project's defer date
					
					if thisProject's defer date is not missing value and thisProject's defer date ³ currentDate then
						
						set projectStatus to "deferred (project)"
						
					end if
					
					--- rename project's status from "active status" to "active"
					
					if projectStatus is "active status" then
						
						set projectStatus to "active"
						
					end if
					
					--- project is on hold
					
				else if thisProject's status is on hold status then
					
					set projectStatus to "on hold"
					
				end if
				
			end if
			
			set text of third cell of newRow to projectStatus
			
			(*
			
			add OmniFocus project due date
			
			*)
			
			if due date of thisProject is not equal to missing value then
				
				set value of fourth cell of newRow to due date of thisProject as date
				
			end if
			
			(*
			
			add OmniFocus root folder name of the project
			
			*)
			
			set text of fifth cell of newRow to rootFolderName as string
			
		end repeat
		
	end tell
	
end tell
