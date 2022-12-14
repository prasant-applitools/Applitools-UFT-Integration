Function eyes_factory( current_page_reference)
	Set eyes = New Applitools_Eyes
	eyes.set_test_name = Environment("TestName")
	Set eyes.set_page_reference = current_page_reference
	eyes.set_extension_installed_status()
	Set eyes_factory = eyes
End Function

Class Check
	Private name
	Private is_full_screen
	Private match_level
	Private ignore_caret
	Private ignore_regions
	Private ignore_displacements
	
End Class

Class Applitools_Eyes

	Private runner_type
	Private local_view_port
	Private page_object
	Private test_name
	Private app_name 
	Private api_key
	Private eye_open_timeout
	Private is_eye_opened
	Private is_extension_installed
	Private concurrency
	Private eyes_server_url
	Private browser_config
	Private batch_name
	Private batch_id
	Private layout_breakpoint
	
	Public Sub set_extension_installed_status()
		check_extension_installed_cmd = "typeof(__applitools)"
		result = execute_command(check_extension_installed_cmd)
		
		If result="undefined" Then
			is_extension_installed = False
		Else
			is_extension_installed = True
		End If
	End Sub
	
	Private Sub Class_Initialize()
		' Read variables from Environment File		
		runner_type = Environment.Value("Applitools_RUNNER")
		app_name = Environment.Value("Applitools_APP_NAME")
		api_key = Environment.Value("Applitools_API_KEY")
		eye_open_timeout = Environment.Value("Applitools_EYE_OPEN_TIMEOUT")
		concurrency = Environment.Value("Applitools_MAX_CONCURRENCY")
		eyes_server_url = Environment.Value("Applitools_EYES_SERVER_URL")
		local_view_port = Environment.Value("Applitools_LOCAL_VIEW_PORT")
		Err.clear
		On Error resume Next
		b_config = Environment.Value("Applitools_BROWSER_CONFIG")
		browser_config = ",browsersInfo:" + b_config
		If Err.number <> 0 Then
			browser_config = ""
		End If
		
		On Error Goto 0
		Err.clear
		
		batch_name = Environment.Value("Applitools_BATCH_NAME")
		batch_id = Environment.Value("Applitools_BATCH_ID")
		layout_breakpoint=Environment.Value("Applitools_BREAKPOINT_LAYOUT")
	End  Sub
	
	Public Property Let set_eye_open_timeout(eye_open_timeout_In_msec)
		eye_open_timeout=eye_open_timeout_In_msec
	End Property
	
	Public Property Set set_page_reference(page_reference)
		Set page_object = page_reference
	End Property
	
	Public Property Let set_test_name(test_name_str)
		test_name = test_name_str
	End Property
	
	Public Property Let set_app_name(app_name_str)
		app_name = app_name_str
	End Property
	
	Public Property Let set_api_key(api_key_str)
		api_key = api_key_str
	End Property
	
	Private Function execute_command(command)
		execute_command = page_object.RunScript(command)
	End Function
	
	Private Function wait_for_check(check_name)
		MercuryTimers(check_name).Start
		
		get_check_status_cmd="__applitools.check.status"
		
		is_success = False
		
		Do While MercuryTimers(check_name).ElapsedTime < eye_open_timeout
		
			check_status = execute_command(get_check_status_cmd)
			Select Case check_status
			Case "SUCCESS"
				is_success = True
				MercuryTimers(check_name).Stop
				Exit Do
			Case "ERROR"
				MercuryTimers(check_name).Stop
				is_success = False
				Exit Do
			Case else
				MercuryTimers(check_name).Continue
			End Select
			
			Wait 1
		Loop
		
		Set check_result = CreateObject("Scripting.Dictionary")
		
		If MercuryTimers(check_name).ElapsedTime > eye_open_timeout Then
			check_result.add "result", False
			check_result.add "message", "Check Operation Timed Out"
		Else
			If is_success Then
				check_result.add "result", True
				check_result.add "message", "Check Operation Success"
			Else
				check_result.add "result", False
				get_check_error_cmd="__applitools.check.error"
				check_result.add "message", execute_command(get_check_error_cmd)
			End If
		End If
		
		Set wait_for_check = check_result
		
	End Function
	
	Private Function wait_for_eye_to_open()
	
		get_open_eye_status_cmd="__applitools.result.status"
		
		Set result = CreateObject("Scripting.Dictionary")
		
		Do While MercuryTimers(test_name).ElapsedTime < eye_open_timeout
			eye_status = execute_command(get_open_eye_status_cmd)
			
			Select Case eye_status
			Case "OPEN"
				is_eye_opened = True
				MercuryTimers(test_name).Stop
				Exit Do
			Case "ERROR"
				is_eye_opened = False
				MercuryTimers(test_name).Stop
				Exit Do
			Case else
				MercuryTimers(test_name).Continue
			End Select
			
			Wait 1
		Loop
		
		If MercuryTimers(test_name).ElapsedTime > eye_open_timeout Then
			result.Add "result", False
			result.Add "message", "Eyes Open timed out" 
			is_eye_opened = False
		Else
			If is_eye_opened Then
				result.Add "result", True
				result.Add "message", "Eyes Opened" 
			Else
				result.Add "result", False
				result.Add "message", "Eyes Opened Errored"
			End If
		End If
		
		set wait_for_eye_to_open = result
	End Function
	
	Public Function open_eye()
	
		If is_extension_installed Then

			batch_info = ",batch:{name:'"  + batch_name + "', id:'" + batch_id + "'}"
			
			eyes_config = ",config: {defaultMatchSettings:{useDom:true, enablePatterns:true}, sendDom:true, appName: '" + app_name + "',testName: '" + test_name + "',apiKey: '" + api_key + "',viewportSize: " + local_view_port + ",serverUrl:'" + eyes_server_url + "'" + browser_config + batch_info + "}"
			
			eyes_open_command ="__applitools.openEyes({type:'" + runner_type + "',stitchMode:'css',layoutBreakpoints:" + layout_breakpoint + ", concurrency: " + concurrency + eyes_config + "}).then(value =>{console.log('Open Completed'); __applitools.result = {status: 'OPEN', value}; console.log(__applitools.result.status)}).catch(error =>{ __applitools.result = {status: 'ERROR', error:error.message};console.log(error.message)}); "
			
			execute_command(eyes_open_command)
			
			Set open_eye = wait_for_eye_to_open()
		Else
			Set result = CreateObject("Scripting.Dictionary")
			result.Add "result", False
			result.Add "message", "Applitools Extension not installed"
			Set open_eye = result
		End If
		
	End Function
	
	Public Function check_full_screen(check_name,algorithm)
	
'		reset_status "check","NEW"
		If is_eye_opened Then
			eyes_check_full_screen_cmd = "console.log('Check : " + check_name + "');__applitools.eyes.check({settings: {name: '" + check_name + "', fully: true, matchLevel: '" + algorithm + "'}}).then(value => {__applitools.check = {status: 'SUCCESS', value}}).catch(error =>{ __applitools.check = {status: 'ERROR', error:error.message}});"
			page_object.RunScript(eyes_check_full_screen_cmd)
			Set check_full_screen = wait_for_check(check_name)
		Else
			Set check_result = CreateObject("Scripting.Dictionary")
			check_result.Add "result", False
			check_result.Add "message", "Eye Not Open"
			Set check_full_screen = check_result
		End If
		
		reset_status "check","COMPLETED"
		
	End Function
	
	Private Sub reset_status(variable_name,expected_status)
		reset_eye_open_result = "__applitools." + variable_name + "={status: '" + expected_status + "',value:'" + expected_status + "'}"
		execute_command(reset_eye_open_result)
	End Sub
	
	Private Sub Class_Terminate()
	
		If is_eye_opened Then
			eyes_close_cmd = "(async function(){console.log('Close'); return await __applitools.eyes.close()})()"
			execute_command(eyes_close_cmd)	
		End If
		
		reset_status "result","COMPLETED"
		
	End  Sub
	
End Class
