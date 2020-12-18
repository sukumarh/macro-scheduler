# Macro Scheduler
Powershell scripts to schedule multiple macros through Windows Task Scheduler.


### Required
1. Windows 10
2. Powershell
3. [Windows Task Scheduler](https://docs.microsoft.com/en-us/windows/win32/taskschd/task-scheduler-start-page)


## Usage
The `action.ps1` script is the entry script to the functionality. This can be triggered manually using PowerShell or scheduled using Windows Task Scheduler.

### Run using PowerShell
The following command can be used to trigger the powershell script from command prompt.
```sh
  powershell -executionpolicy bypass -file C:\directory-to-repo\macro-scheduler\action.ps1
```

### Run using Task Scheduler
Create a new task in Task Scheduler and add an Action with the following parameters.
| Parameter      | Value                                                                         |
| -------------- | ----------------------------------------------------------------------------- |
| Action         | Start a program                                                               |
| Program/Script | powershell                                                                    |
| Add arguments  | -executionpolicy bypass -file C:\directory-to-repo\macro-scheduler\action.ps1 |


## Configuration
In the `config.xml`, add the files that need to be triggered sequentially along with the macros in each file. Provide a label for each file, for the title of the toast notification.

#### Notification level
1. Macros
    ```xml
      <NotifyFor>Macros</NotifyFor>
    ```

2. Files
    ```xml
      <NotifyFor>Files</NotifyFor>
    ```

#### Notification message
The message to be displayed in the toast notification. 
```xml
  <NotificationMessage>Macro executed.</NotificationMessage>
```

If the notification level is _Macros_, the macro name can be passed into the message using `{0}`.  
```xml
  <NotificationMessage>{0} executed.</NotificationMessage>
```
