import System
import System.Management.Automation
import System.Management.Automation.Runspaces
import System.Text
import System.Reflection
import System.Collections.ObjectModel

// Function to load the PowerShell assembly dynamically
def LoadAssembly():
    try:
        Console.ForegroundColor = ConsoleColor.Yellow
        Console.WriteLine("[*] Loading assemblies")
        Console.ResetColor()
        
        asm = Assembly.Load("System.Management.Automation, Version=3.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35")
        if asm is not null:
            Console.ForegroundColor = ConsoleColor.Green
            Console.WriteLine("[+] Assembly loaded successfully")
            Console.ResetColor()
            return true
    except ex as Exception:
        Console.ForegroundColor = ConsoleColor.Red
        Console.WriteLine("[-] Error loading assembly: ${ex.Message}")
        Console.ResetColor()
        return false

// Function to create and open a PowerShell runspace
def CreateRunspace():
    try:
        iss = InitialSessionState.CreateDefault()
        runspace = RunspaceFactory.CreateRunspace(iss)
        runspace.Open()
        return runspace
    except ex as UnauthorizedAccessException:
        Console.ForegroundColor = ConsoleColor.Red
        Console.WriteLine("[-] Unauthorized Access: ${ex.Message}")
        Console.ResetColor()
        return null
    except ex as Exception:
        Console.ForegroundColor = ConsoleColor.Red
        Console.WriteLine("[-] Error creating runspace: ${ex.Message}")
        Console.ResetColor()
        return null

// Function to execute a PowerShell command and return results
def ExecuteCommand(runspace as Runspace, command as string) as Collection[PSObject]:
    using pipeline = runspace.CreatePipeline():
        pipeline.Commands.AddScript(command)
        try:
            return pipeline.Invoke()
        except ex as Exception:
            Console.ForegroundColor = ConsoleColor.Red
            Console.WriteLine("[-] Error executing command: ${ex.Message}")
            Console.ResetColor()
            return null

// Main function to start the PowerShell session and interact with the user
def Main():
    if not LoadAssembly():
        return
    
    runspace = CreateRunspace()
    if runspace is null:
        return
    
    Console.ForegroundColor = ConsoleColor.Yellow
    Console.WriteLine("[*] Starting PowerShell Session")
    Console.ResetColor()
    
    while true:
        Console.ForegroundColor = ConsoleColor.Cyan
        Console.Write("PS > ")
        Console.ResetColor()
        
        command = Console.ReadLine()
        
        if string.IsNullOrEmpty(command):
            continue
            
        if command.ToLower() == "exit":
            break
        
        results = ExecuteCommand(runspace, command)
        if results is not null:
            for result in results:
                if result is not null:
                    Console.WriteLine(result.ToString())
    
    runspace.Close()
    Console.ForegroundColor = ConsoleColor.Yellow
    Console.WriteLine("[*] Session terminated")
    Console.ResetColor()

// Entry point
Main()
