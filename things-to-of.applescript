--------------------------------------------------
--------------------------------------------------
-- Import tasks from Things to OmniFocus
--------------------------------------------------
--------------------------------------------------
--
-- Script taken from: http://forums.omnigroup.com/showthread.php?t=14846&page=2
-- Added: creation date, due date, start date functionality
-- Empty your Things Trash first.
-- Note that this won't move over scheduled recurring tasks.

set progress total steps to 0
set progress completed steps to 0
set progress description to ""
set progress additional description to ""

tell application "Things"
	set listOfTodos to to dos
end tell

set numberOfTodos to count my listOfTodos
set progress total steps to numberOfTodos
set progress completed steps to 0
set progress description to "Processing Todos..."
set progress additional description to "Preparing to process."
set currentPosition to 0

repeat with aToDo in listOfTodos
	set currentPosition to currentPosition + 1
	
	set progress additional description to "Processing Todo " & currentPosition & " of " & numberOfTodos
	
	tell application "Things"
		-- Get title and notes of Things task
		set theTitle to name of aToDo
		set theNote to notes of aToDo
		set theCreationDate to creation date of aToDo
		set theDueDate to due date of aToDo
		set theStartDate to activation date of aToDo
		set toDoCompletion to completion date of aToDo
		
		-- get dates
		if (creation date of aToDo) is not missing value then
			set theCreationDate to creation date of aToDo
		else
			set theCreationDate to current date
		end if
		
		if (due date of aToDo) is not missing value then
			set theDueDate to due date of aToDo
		end if
		
		if (activation date of aToDo) is not missing value then
			set theStartDate to activation date of aToDo
		end if
		
		-- Get project name
		if (project of aToDo) is not missing value then
			set theProjectName to (name of project of aToDo)
		else
			set theProjectName to "NoProjInThings"
		end if
		
		-- Get Contexts from tags
		-- get all tags from one ToDo item...
		set allTagNames to {}
		copy (name of tags of aToDo) to allTagNames
		-- ...and from the project...
		if (project of aToDo) is not missing value then
			copy (name of tags of project of aToDo) to the end of allTagNames
			
			-- ...and from the area...
			if (area of project of aToDo) is not missing value then
				copy (name of tags of area of project of aToDo) to the end of allTagNames
			end if
		end if
		
		-- ...now extract contexts from tags
		copy my FindContextName(allTagNames) to theContextName
		
		-- Create a new task in OmniFocus
		tell application "OmniFocus"
			
			tell default document
				
				-- Set (or create new) task context
				if context theContextName exists then
					set theContext to context theContextName
				else
					set theContext to make new context with properties {name:theContextName}
				end if
				
				-- Set (or create new) project
				if project theProjectName exists then
					set theProject to project theProjectName
				else
					set theProject to make new project with properties {name:theProjectName, singleton action holder:true}
				end if
				
				-- Create new task
				tell theProject
					set newTask to make new task with properties {name:theTitle, note:theNote, context:theContext, creation date:theCreationDate}
					
					
					if (theStartDate is not missing value) then set the defer date of newTask to theStartDate
					
					if (theDueDate is not missing value) then set the due date of newTask to theDueDate
					
					if (toDoCompletion is not missing value) then set the completion date of newTask to toDoCompletion
					
					if (toDoCompletion is not missing value) then set the completed to true
					
				end tell
			end tell -- document
		end tell -- OF application
	end tell
end repeat
set progress total steps to 0
set progress completed steps to 0
set progress description to ""
set progress additional description to ""

--------------------------------------------------
--------------------------------------------------
-- Get context from array of Things tags
--------------------------------------------------
--------------------------------------------------
on FindContextName(tagNames)
	repeat with aTagName in tagNames
		return aTagName
	end repeat
	return ""	
end FindContextName