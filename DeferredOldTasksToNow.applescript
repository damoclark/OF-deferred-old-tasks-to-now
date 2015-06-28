(*
	This script scans all projects and action groups in the front OmniFocus document identifying any that
	have a defer date prior to today, and updates their defer date to today so they remain in the Forecast Perspective view.
	
	by Damien Clark based on code written by Curt Clifton
	
	Copyright © 2007-2014, Curtis Clifton
	Copyright © 2014, Damien Clark
	All rights reserved.
	
	Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:
	
		¥ Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.
		
		¥ Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.
		
	THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
	
	version 0.1: Original release
*)

(*
	The following property is used for script notifications.
*)
property scriptSuiteName : "DeferOldTasksToNow"

--Here we go!
tell application "OmniFocus"
	tell front document
		try
			set will autosave to false
			set deferredTaskCount to my updateDeferDates(it)
			log "Number of deferred tasks updated to today: " & deferredTaskCount
			set will autosave to true
		on error errText number errNum
			set will autosave to true
			display dialog "Error: " & errNum & return & errText
			return
		end try
		my notify("Remaining Deferred Tasks", "" & deferredTaskCount & " old tasks deferred to today")
	end tell
end tell

(* 
	Sets the defer date to today for all incomplete tasks with an existing defer date in all active projects found in the container document
	theContainer: a document or folder containing flattened projects
*)
on updateDeferDates(theContainer)
	log theContainer
	using terms from application "OmniFocus"
		set theProjects to every flattened project of theContainer whose status is active
		set deferredTaskCount to my updateDeferDatesProjects(theProjects)
	end using terms from
	return deferredTaskCount
end updateDeferDates

(* 
	Recurses over the trees rooted at the given projects, identifying items that are:
		¥ not complete and 
		¥ have a defer date, prior to today
		and sets their defer date to today
	theProjects: a list of projects
*)
on updateDeferDatesProjects(theProjects)
	set deferredTaskCount to 0
	using terms from application "OmniFocus"
		repeat with aProject in theProjects
			log "Checking project: " & (name of aProject)
			set projectContainer to container of aProject
			if (class of projectContainer is not folder or (class of projectContainer is folder and projectContainer is not hidden)) then
				set theRootTask to root task of aProject
				set deferredTaskCount to deferredTaskCount + (my updateDeferDatesTask(theRootTask, true))
			else
				log "skipped" & (name of aProject)
			end if
		end repeat
	end using terms from
	return deferredTaskCount
end updateDeferDatesProjects

(* 
	Recurses over the tree rooted at the given task, looking for items that are:
		¥ not complete and 
		¥ have a defer date, prior to today
		and sets their defer date to today
	theTask: a task
	isProjectRoot: true iff theTask is the root task of a project
*)
on updateDeferDatesTask(theTask, isProjectRoot)
	set deferredTaskCount to 0
	using terms from application "OmniFocus"
		tell theTask
			--If it is not completed and has a defer date less than today
			if (completed is false and defer date is not missing value and defer date is less than (current date)) then
				--set the defer date to the current date
				log "Setting defer date of task " & name & " from " & defer date & " to " & (current date)
				set defer date to current date
				set deferredTaskCount to deferredTaskCount + 1
			end if
			
			--If theTask is completed or is does not have any subtasks, then return
			set isAProjectOrSubprojectTask to isProjectRoot or (count of (get tasks)) > 0
			if (completed or not isAProjectOrSubprojectTask) then return deferredTaskCount
			
			--Otherwise, get all sub-tasks that arent completed
			set incompleteChildTasks to every task whose completed is false
			--If there are some
			if ((count incompleteChildTasks) is not 0) then
				--Then update their defer dates too
				repeat with aTask in incompleteChildTasks
					set deferredTaskCount to deferredTaskCount + (my updateDeferDatesTask(aTask, false))
				end repeat
			end if
		end tell
	end using terms from
	return deferredTaskCount
end updateDeferDatesTask

(*
	Uses Notification Center to display a notification message.
	theTitle Ð a string giving the notification title
	theDescription Ð a string describing the notification event
*)
on notify(theTitle, theDescription)
	display notification theDescription with title scriptSuiteName subtitle theTitle
end notify
