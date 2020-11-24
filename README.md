# ProcessOutputMonitor 
It was hard to find in the [public domain](https://stackru.com/questions/5223319/hta-kak-graficheskij-interfejs-komandnoj-stroki), but not on GitHub, so I uploaded it.

# What it do?
`ProcessOutputMonitor` can spawns a child process and monitors the output.

# How to use?
#### Public Functions
`.PID`                - Return child process `Process ID`.   
`.IsRunning`          - Checks the child process status. Returns boolean.   
`.NextLine`           - Return output per line    
`.NextData`           - Return new output

#### 1. Declare ProcessOutputMonitor object
```
Dim monitorGoogle
Set monitorGoogle = New ProcessOutputMonitor
```
#### 2. Starts the child process with 
```
monitorGoogle.Start "cmd.exe /c ping -4 www.google.com"
```
#### 3. Start the Sub that monitor the child process in 500ms interval
```
timerID = window.setInterval(GetRef("monitorPings"), 500)
```
#### In the Monitor Sub, checks the child process output
```
Dim buffer, keepRunning
keepRunning = False

buffer = monitorGoogle.NextData
If Not IsEmpty( buffer ) Then 
    ' This will run if the child process shows output at cmd
    ' Put your codes here
    
    keepRunning = True
Else 
    ' This will run if the child process is still running and output is empty
    ' Put your codes here
    
    keepRunning = CBool( keepRunning Or monitorGoogle.IsRunning )
End If 
```
#### In Monitor Sub, check if child process has stopped and is not running
```
If Not keepRunning Then 
    window.clearInterval( timerID )
    timerID = Empty
    ' This will run once the child process has stopped running
    ' Put your codes here
End If 
```





