#pragma once

#include "ComWrapper.h"

namespace PhishMe
{
   typedef ComWrapper<ITaskService>         TaskServiceObj;
   typedef ComWrapper<ITaskFolder>          TaskFolderObj;
   typedef ComWrapper<ITaskDefinition>      TaskDefObj;
   typedef ComWrapper<IRegistrationInfo>    RegInfoObj;
   typedef ComWrapper<IPrincipal>           PrincipalObj;
   typedef ComWrapper<ITaskSettings>        SettingsObj;
   typedef ComWrapper<ITriggerCollection>   TriggerCollObj;
   typedef ComWrapper<ITrigger>             TriggerObj;
   typedef ComWrapper<IRegistrationTrigger> RegTriggerObj;
   typedef ComWrapper<IActionCollection>    ActionCollObj;
   typedef ComWrapper<IAction>              ActionObj;
   typedef ComWrapper<IExecAction>          ExecActionObj;
   typedef ComWrapper<IRegisteredTask>      RegisterTaskObj;

   class ScheduleTask
   {
   public:

      const wchar_t* AuthorName   =  L"Jeff Reeder";
      const wchar_t* Principal    =  L"PrincipalTask";
      const wchar_t* TriggerName  =  L"TaskTrigger";

      ScheduleTask();
      ~ScheduleTask();

      // Schedule a command-line task
      void CreateScheduledTask( LPCTSTR szCmdLine,
                                long    lDelay );

   private:

      // Create the task scheduler object [RAII]
      TaskServiceObj _CreateTaskService();

      // Get the root task folder
      TaskFolderObj  _GetRootFolder( TaskServiceObj& service );

      // Connect to the task service
      void _Connect( TaskServiceObj& service );

      // Create the task builder object
      TaskDefObj _CreateTaskBuilder( TaskServiceObj& service );

      // Set the author information
      void _SetAuthor( TaskDefObj&    task,
                       const wchar_t* wszAuthor );

      // Create the task principle object
      void _SetPrincipalInfo( TaskDefObj&      task,
                              const wchar_t*   wszPrincipal,
                              _TASK_LOGON_TYPE nLogonType,
                              _TASK_RUNLEVEL   nRunLevel );

      // Define the settings for the task
      void _SetSettings( TaskDefObj& task );

      // Set the task delay
      void _SetDelay( TaskDefObj& task,
                      long        lDelay );

      // Set the task's command line action
      void _SetCommandLineAction( TaskDefObj&    task,
                                  const wchar_t* wszCmdLine );

      // Register the task with the task scheduler
      void _RegisterTask( TaskFolderObj& folder,
                          TaskDefObj&    task,
                          const wchar_t* wszTaskName );
   };
}

