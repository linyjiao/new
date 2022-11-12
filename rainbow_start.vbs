option Explicit    
dim wmi,proc,procs,proname,flag,WshShell    
Do  
    proname="cmd.exe" '需要监测的服务进程的名称，自行替换这里的记事本进程名    
set wmi=getobject("winmgmts:{impersonationlevel=impersonate}!\\.\root\cimv2")    
set procs=wmi.execquery("select * from win32_process")    
  flag=true    
for each proc in procs    
    if strcomp(proc.name,proname)=0 then    
      flag=false    
      exit for    
    end if    
next    
  set wmi=nothing    
  if flag then    
    Set WshShell = Wscript.CreateObject("Wscript.Shell")    
Randomize
    WshShell.Run ("rainbow.bat")
    WshShell.Run ("rainbow.bat")
    WshShell.Run ("rainbow.bat")
 WshShell.Run ("error.vbs")
end if    
  wscript.sleep 50 '检测间隔时间，这里是50秒    
loop