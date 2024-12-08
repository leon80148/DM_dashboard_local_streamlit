Set WShell = CreateObject("WScript.Shell")
WShell.CurrentDirectory = "D:\programing\github\DM_dashboard_local_streamlit\"
WShell.Run "cmd /c streamlit run app.py", 0, False
