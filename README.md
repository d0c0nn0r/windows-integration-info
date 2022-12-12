# Introduction

This project is for a console application, that can be used on Legacy / Modern systems to gather critical hardware, software and operating system details  
needed to assess system changes potentially required to integrate that system into OSI-PI.

## System Prerequisites

Program compiles to .NET 3.5. Any system where this console app (.exe) will be run must have .NET Framework 3.5 or greater installed, otherwise  
the app will not run.
No other auxiliary dependencies exist.

## Build and Test

Build will produce 2 artifacts on every build:

1. windevops.cli.exe

    - This is a compiled executable, with all auxiliary dll's merged so that the entire program can be delivered as a single .exe.
    - This .exe is portable, and can be run directly from the cmdline on any machine without any installation (provided .NET 3.5 is installed)

2. Release.zip

- This is a zip of all the application build files, for posterity.
- Only the windevops.cli.exe should be used

## Use

The command line tool help can be accessed as shown in the below gif.

![cli gif](./Images/windevops.cli.gif)

## What it does

This commandline application gathers information from a computer system, and stores that to csv, txt or other file types.  
The information is gathers is;

- **Hardware details:** 

Everything about CPU, Disks, Network Cards, Operating System, BIOS and chassis

- **Software details:** 

All software installed, including version, publisher, install location, date installed, and other miscellaneous details.

- **Process details:** 

All processes running on the local machine

- **Local Windows Security:** 

All Local Windows Groups, and the members in them

- **Service details:** 

All windows services, their startup type and the Logon User they run under

- **Time details:** 

All clock information, such as Timezone and Time Sync source

- **Firewall Rule details:** 

All inbound and outbound public, private and domain firewall rules

- **DCOM details:** 

All Global Computer DCOM settings, and each and every DCOM Application settings as well. Includes security permissions, options and service details

- **Local Security Policy:** 

Local Audit, User Rights and Security Options setting and security permissions applied.

- **MS SQL Server details:** 

All application settings, such as Ports, Instance Name, Version

- **Oracle details:** 

Captures the local tnsnames.ora file, if it exists

- **NetStat Table:** 

Captures the INBOUND and OUTBOUND network connections and the corresponding process ID that is creating the connections

- **Kepware/RSLinx Projects:** 

Captures the local kepware and/or RSLinx project files if they exist

- **AutoLogin details:** 

The registry key relating to how AutoLogin is configured

- **ODBC details:** 

The registry key, and all subkeys, relating to local ODBC DNS settings


## Information Gathering

A pre-built example [batch file](./gatherLocal.bat) is provided.

- Copy the windevops.cli.exe to any machine, 
- Copy or manually recreate this [batch file](./gatherLocal.bat).
- Open a cmd shell, and execute the batch file
- Copy all files created in the "C:\temp\piInfoGathering" 