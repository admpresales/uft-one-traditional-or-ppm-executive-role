'===========================================================================================
'20201007 - DJ: Initial creation
'20210115 - DJ: Disabled smart identification
'===========================================================================================

Dim BrowserExecutable, Counter

While Browser("CreationTime:=0").Exist(0)   												'Loop to close all open browsers
	Browser("CreationTime:=0").Close 
Wend
BrowserExecutable = DataTable.Value("BrowserName") & ".exe"
SystemUtil.Run BrowserExecutable,"","","",3													'launch the browser specified in the data table
Set AppContext=Browser("CreationTime:=0")													'Set the variable for what application (in this case the browser) we are acting upon

'===========================================================================================
'BP:  Navigate to the PPM Launch Pages
'===========================================================================================

AppContext.ClearCache																		'Clear the browser cache to ensure you're getting the latest forms from the application
AppContext.Navigate DataTable.Value("URL")													'Navigate to the application URL
AppContext.Maximize																			'Maximize the application to give the best chance that the fields will be visible on the screen
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Click the Executive Overview link
'===========================================================================================
Browser("Browser").Page("Project & Portfolio Management").Image("Executive Overview Image Link").Click
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Click the Ron Steel (CIO) link to launch PPM as Ron Steel
'===========================================================================================
Browser("Browser").Page("PPM Launch Page").WebArea("Ron Steel Image Link").Click
AppContext.Sync	

If Browser("Browser").Page("Dashboard - Overview Dashboard").WebElement("Size of bubble text").Exist Then
	Reporter.ReportEvent micPass, "Find the bubble text", "The text did display within the default .Exist timeout, this means that the Porfolio Dashboard portlet didn't load"
Else 
	Reporter.ReportEvent micFail, "Find the bubble text", "The text didn't display within the default .Exist timeout, this means that the Porfolio Dashboard portlet didn't load"
End If

'===========================================================================================
'BP:  Hover over each Business Objective category to capture the changes in the Porfolio Scorecard
'===========================================================================================
Browser("Browser").Page("Dashboard - Overview Dashboard").WebElement("Regulatory Compliance").HoverTap
Browser("Browser").Page("Dashboard - Overview Dashboard").WebElement("9 Month Release Cycle").HoverTap
Browser("Browser").Page("Dashboard - Overview Dashboard").WebElement("Reduce Customer Churn").HoverTap
Browser("Browser").Page("Dashboard - Overview Dashboard").WebElement("10% Increase in Revenue").HoverTap
Browser("Browser").Page("Dashboard - Overview Dashboard").WebElement("15% Growth in Partner").HoverTap
Browser("Browser").Page("Dashboard - Overview Dashboard").WebElement("Cost Containment").HoverTap

'===========================================================================================
'BP:  Verify that the Budget by Business Objective dashboard element is displayed
'		Added a traditional OR click on the dashboard name to force scroll if the 
'		resolution of the machine is too small to have the hamburger menu be displayed
'===========================================================================================
Browser("Browser").Page("Dashboard - Overview Dashboard").WebElement("Budget by Business Objective (This Year)").Click
Browser("Browser").Page("Dashboard - Overview Dashboard").Image("Budget by Business Objective Hamburger Menu").Click
AppContext.Sync																				'Wait for the browser to stop spinning
Browser("Browser").Page("Dashboard - Overview Dashboard").Link("Maximize").Click
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Hover over each Business Objective category to capture the changes in the Porfolio Scorecard
'===========================================================================================
Browser("Browser").Page("Maximized Portlet: Budget").WebElement("Regulatory Compliance").HoverTap
Browser("Browser").Page("Maximized Portlet: Budget").WebElement("Reduce Customer Churn").HoverTap
Browser("Browser").Page("Maximized Portlet: Budget").WebElement("Cost Containment").HoverTap
Browser("Browser").Page("Maximized Portlet: Budget").WebElement("9 Month Release Cycle").HoverTap
Browser("Browser").Page("Maximized Portlet: Budget").WebElement("15% Growth in Partner").HoverTap
Browser("Browser").Page("Maximized Portlet: Budget").WebElement("10% Increase in Revenue").HoverTap

'===========================================================================================
'BP:  Search for Open the Portfolio (ITFM) Dashboard
'===========================================================================================
Browser("Browser").Page("Dashboard - Overview Dashboard").Link("DASHBOARD").Click
Browser("Browser").Page("Dashboard - Overview Dashboard").Link("Private").HoverTap
Browser("Browser").Page("Dashboard - Overview Dashboard").Link("ITFM").HoverTap
Browser("Browser").Page("Dashboard - Overview Dashboard").Link("Portfolio (ITFM)").Click
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Click the Trial Portfolio to exercise drill down
'===========================================================================================
Browser("Browser").Page("Dashboard - Overview Dashboard").WebElement("Trial Portfolio").Click
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Click the Marketing WebPortaI V2 to exercise drill down to the project dashboard
'===========================================================================================
Browser("Browser").Page("Cost Plan").Link("Marketing WebPortal V2").Click
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Click the down triangle to show you could override the calculated health
'===========================================================================================
Browser("Browser").Page("Project Overview").WebElement("Down Triangle").Click
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Click Done button
'===========================================================================================
Browser("Browser").Page("Project Overview").Frame("overrideProjectHealthDialogIF").WebButton("Done").Click
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Logout.  Use traditional OR
'===========================================================================================
Browser("Browser").Page("Project Overview").WebElement("menuUserIcon").Click
AppContext.Sync																				'Wait for the browser to stop spinning
Browser("Browser").Page("Project Overview").Link("Sign Out Link").Click
AppContext.Sync																				'Wait for the browser to stop spinning

AppContext.Close																			'Close the application at the end of your script

