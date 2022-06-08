# VisualStudioFileOpenTool
Visual Studio File Open Tool

	usage: <openf|debug|openp|stopdebug|build|rebuild> <visual_studio_path> <solution_path> [<file_path> <file_line_number>] |[<PID_of_process_to_attach_to_debugger> | <Build_Config>]
	
	The arguments in the bracket are conditional. If first agument is "openf", then filePath and fileLine is required, if it is "debug" PID is required otherwise no fourth argument is needed.
	
	<openf> is for opening file in a solution
	
	<openp> is for simply openening the solution
	
	<debug> is to attach a process to the Visual Studio debugger and start debugging that process
	
	<stopdebug> is for stopping the debugger and detaching any attached process
	
	<build> is for building the solution
	
	<rebuild> is for rebuilding the solution
	
	<visual_studio_path> is the path to visual studio executable(devenv.exe)
	
	<solution_path> is the path to the project solution(.sln) file
	
	<file_path> is the path to the file in the project that needs to be opened
	
	<file_line_number> is the line number that should be in focus after opening the file at <file_path>
	
	<PID_of_process_to_attach_to_debugger> is the process ID of the process that needs debugging by attaching it to the Visual Studio Debugger
	
	<Build_Config> The config of the build to made(Debug, Release)