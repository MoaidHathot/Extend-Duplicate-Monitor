shell = WScript.CreateObject("WScript.Shell");

editor = shell.Exec("%systemroot%\\system32\\control.exe desk.cpl,@0,3");

shell.AppActivate( editor.ProcessID );

WScript.Sleep(500);

shell.SendKeys("{TAB}");
shell.SendKeys("{TAB}");
shell.SendKeys("{TAB}");
shell.SendKeys("{TAB}");

shell.SendKeys("%{DOWN}");
//shell.SendKeys("{UP}");
shell.SendKeys("d");
shell.SendKeys("{ENTER}");

WScript.Sleep(2000);

shell.SendKeys("{TAB}");
shell.SendKeys("{ENTER}");