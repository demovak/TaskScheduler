#include "stdafx.h"

#include "ScheduleTask.h"
#include "com_exception.h"
#include "ComWrapper.h"


using namespace std;
using namespace PhishMe;


ScheduleTask::ScheduleTask()
{
}


ScheduleTask::~ScheduleTask()
{
}


// Schedule a command-line task
void ScheduleTask::CreateScheduledTask( LPCTSTR szCmdLine,
                                        long    lDelay )
{
   // TODO: Validate szCmdLine and lDelay

   //  Create a name for the task.
   LPCWSTR wszTaskName = L"PhishMe Task Scheduler Sample";

   TaskServiceObj taskService  =  _CreateTaskService();        // Create the task service

   _Connect(taskService);                                      // Connect to the task service

   TaskFolderObj rootFolder  =  _GetRootFolder(taskService);   // Get the root folder

   // If the same task exists, remove it - we ignore the return value,
   // because it can legitimately fail if the task doesn't exist
   HRESULT hr  =  rootFolder.GetObj()->DeleteTask( _bstr_t(wszTaskName), 0 );

   // Create the task builder object to create task
   TaskDefObj task  =  _CreateTaskBuilder(taskService);

   // TODO: No longer need taskService - make it auto-release early

   // Set the task's author name
   _SetAuthor( task, AuthorName );

   // Create the task principal info
   _SetPrincipalInfo( task,
                      AuthorName,
                      TASK_LOGON_INTERACTIVE_TOKEN,
                      TASK_RUNLEVEL_LUA );

   _SetSettings(task);
   _SetDelay( task, lDelay );

   // Set the command line for this task
#ifdef _UNICODE
   _SetCommandLineAction( task, szCmdLine );
#else
   // Convert ANSI string to Unicode
   DWORD   dwMinSize  =  strlen( szCmdLine ) + 1;
   wchar_t wszCmdLine[128];      // WAY bigger than ever possible!  TODO: Handle malloc/free safely

   MultiByteToWideChar( CP_ACP, 0, szCmdLine, -1, wszCmdLine, dwMinSize );

   _SetCommandLineAction( task, wszCmdLine );
#endif

   // Now register the task
   _RegisterTask( rootFolder, 
                  task, 
                  wszTaskName );
}


// Create the task scheduler object [RAII]
TaskServiceObj ScheduleTask::_CreateTaskService()
{
   TaskServiceObj service( "ITaskService" );
   ITaskService*  pService  =  NULL;

   HRESULT hr  =  CoCreateInstance( CLSID_TaskScheduler,
                                    NULL,
                                    CLSCTX_INPROC_SERVER,
                                    IID_ITaskService,
                                    (void**) &pService );

   if ( FAILED(hr) )   throw com_exception( "Failed to create an instance of ITaskService", hr );

   service.SetObj(pService);

   return service;
}


// Connect to the task service
void ScheduleTask::_Connect( TaskServiceObj& service )
{
   HRESULT hr = service.GetObj()->Connect( _variant_t(),
                                           _variant_t(),
                                           _variant_t(),
                                           _variant_t() );

   if ( FAILED( hr ) )   throw com_exception( "ITaskService::Connect failed: %x", hr );
}


// Get the root task folder
TaskFolderObj ScheduleTask::_GetRootFolder( TaskServiceObj& service )
{
   TaskFolderObj rootFolder( "ITaskFolder" );
   ITaskFolder*  pFolder  =  NULL;

   HRESULT hr  =  service.GetObj()->GetFolder( _bstr_t( L"\\" ),
                                               & pFolder );


   if ( FAILED( hr ) )   throw com_exception( "Failed to get task foldercreate an instance of ITaskService", hr );

   rootFolder.SetObj(pFolder);

   return rootFolder;
}


// Create the task builder object
TaskDefObj ScheduleTask::_CreateTaskBuilder( TaskServiceObj& service )
{
   TaskDefObj       taskDef( "ITaskDefinition" );
   ITaskDefinition* pTask  =  NULL;

   HRESULT hr  =  service.GetObj()->NewTask( 0, &pTask );

   if ( FAILED(hr) )   throw com_exception( "Failed to create a task definition", hr );

   taskDef.SetObj(pTask);

   return taskDef;
}


// Set the author information
void ScheduleTask::_SetAuthor( TaskDefObj&    task,
                               const wchar_t* wszAuthor )
{
   if (  wszAuthor == nullptr )   throw runtime_error( "Invalid author parameter in ScheduleTask::_SetAuthor()" );
   if ( *wszAuthor == L'\0'   )   throw runtime_error( "Blank author parameter in ScheduleTask::_SetAuthor()"   );

   RegInfoObj         regInfo( "IRegistrationInfo" );
   IRegistrationInfo* pRegInfo = NULL;

   HRESULT hr = task.GetObj()->get_RegistrationInfo( &pRegInfo );
   if ( FAILED( hr ) )   throw com_exception( "Cannot get identification pointer", hr );

   regInfo.SetObj( pRegInfo );

   hr  =  regInfo.GetObj()->put_Author( _bstr_t(wszAuthor) );
   if ( FAILED(hr) )   throw com_exception( "Cannot set author", hr );

   // All done with need registration - auto-delete
}


// Create the task principle object
void ScheduleTask::_SetPrincipalInfo( TaskDefObj&      task,
                                      const wchar_t*   wszPrincipal,
                                      _TASK_LOGON_TYPE nLogonType,
                                      _TASK_RUNLEVEL   nRunLevel )
{
   if (  wszPrincipal == nullptr )   throw runtime_error( "Invalid principal parameter in ScheduleTask::_SetAuthor()" );
   if ( *wszPrincipal == L'\0'   )   throw runtime_error( "Blank principal parameter in ScheduleTask::_SetAuthor()" );

   PrincipalObj principal( "IPrincipal" );
   IPrincipal*  pPrincipal  =  NULL;

   HRESULT hr  =  task.GetObj()->get_Principal( &pPrincipal );
   if ( FAILED(hr) )   throw com_exception( "Can't create principal for task", hr );

   principal.SetObj(pPrincipal);

   // Set principal ID
   hr  =  principal.GetObj()->put_Id( _bstr_t(wszPrincipal) );
   if ( FAILED(hr) )   throw com_exception( "Can't set principal name for task", hr );

   // Set the logon type
   hr  =  principal.GetObj()->put_LogonType(nLogonType);
   if ( FAILED(hr) )   throw com_exception( "Could not set the principal's logon type", hr );

   //  Run the task with the least privileges (LUA) 
   hr = principal.GetObj()->put_RunLevel(nRunLevel);
   if ( FAILED(hr) )   throw com_exception( "Cannot put principal run level", hr );

   // All done with principal object - auto-releases
}


// Define the settings for the task
void ScheduleTask::_SetSettings( TaskDefObj& task )
{
   SettingsObj    settings( "ITaskSettings" );
   ITaskSettings* pSettings  =  NULL;

   HRESULT hr  =  task.GetObj()->get_Settings( &pSettings );
   if ( FAILED(hr) )   throw com_exception( "Cannot get settings pointer", hr );

   settings.SetObj(pSettings);

   // Set the start when available to TRUE
   hr  =  settings.GetObj()->put_StartWhenAvailable(VARIANT_TRUE);
   if ( FAILED(hr) )   throw com_exception( "Can't set task start when available", hr );

   // Settings object is done - can auto-release
}


// Set the task delay
void ScheduleTask::_SetDelay( TaskDefObj& task,
                              long        lDelay )
{
   if ( lDelay <= 0 )   throw runtime_error( "The task delay must be greater than 0" );

   //  Get the trigger collection to insert the registration trigger.
   TriggerCollObj      triggerColl( "ITriggerCollection" );
   ITriggerCollection* pTriggerCollection  =  NULL;

   HRESULT hr  =  task.GetObj()->get_Triggers( &pTriggerCollection );
   if ( FAILED(hr) )   throw com_exception( "Cannot get trigger collection", hr );

   triggerColl.SetObj(pTriggerCollection);

   //  Add the registration trigger to the task.
   TriggerObj trigger( "ITrigger" );
   ITrigger*  pTrigger = NULL;

   hr  =  triggerColl.GetObj()->Create( TASK_TRIGGER_REGISTRATION,  & pTrigger );
   if ( FAILED(hr) )   throw com_exception( "Cannot create a registration trigger", hr );

   trigger.SetObj(pTrigger);

   RegTriggerObj         regTrigger( "IRegistrationTrigger" );
   IRegistrationTrigger* pRegistrationTrigger  =  NULL;

   hr  =  pTrigger->QueryInterface( IID_IRegistrationTrigger, 
                                    (void**) &pRegistrationTrigger );

   if ( FAILED(hr) )   throw com_exception( "QueryInterface call failed on IRegistrationTrigger", hr );

   regTrigger.SetObj(pRegistrationTrigger);

   hr  =  regTrigger.GetObj()->put_Id( _bstr_t(TriggerName) );
   if ( FAILED( hr ) )   throw com_exception( "Cannot set trigger name", hr );

   // Construct our delay string value
   wstring sDelay  =  wstring( L"PT" )  +  to_wstring(lDelay)  +  L"S";

   //hr = regTrigger.GetObj()->put_Delay( _bstr_t( L"PT30S" ) );
   hr  =  regTrigger.GetObj()->put_Delay( _bstr_t( sDelay.c_str() ) );
   if ( FAILED(hr) )   throw com_exception( "Cannot set registration trigger delay", hr );

   // All done with triggerColl, triggerObj and regTrigger.  All to auto-release
}


// Set the task's command line action
void ScheduleTask::_SetCommandLineAction( TaskDefObj&    task,
                                          const wchar_t* wszCmdLine )
{
   if (  wszCmdLine == nullptr )   throw runtime_error( "Invalid command line parameter in ScheduleTask::_SetCommandLineAction()" );
   if ( *wszCmdLine == L'\0'   )   throw runtime_error( "Blank command line parameter in ScheduleTask::_SetCommandLineAction()" );

   // Get the action collection
   ActionCollObj      actionColl( "IActionCollection" );
   IActionCollection* pActionCollection  =  NULL;

   HRESULT hr  =  task.GetObj()->get_Actions( &pActionCollection );
   if ( FAILED(hr) )   throw com_exception( "Cannot get Task collection pointer", hr );

   actionColl.SetObj(pActionCollection);

   //  Create the action, specifying that it is an executable action.
   ActionObj action( "IAction" );
   IAction*  pAction  =  NULL;

   hr  =  pActionCollection->Create( TASK_ACTION_EXEC, &pAction );
   if ( FAILED( hr ) )   throw com_exception( "Cannot create action", hr );

   action.SetObj(pAction);

   //  Get the execute object
   ExecActionObj execAction( "IExecAction" );
   IExecAction* pExecAction  =  NULL;

   hr  =  pAction->QueryInterface( IID_IExecAction, 
                                   (void**) &pExecAction );

   if ( FAILED(hr) )   throw com_exception( "QueryInterface call failed for IExecAction", hr );
   
   execAction.SetObj(pExecAction);

   //  Set the path of the executable to notepad.exe.
   hr  =  pExecAction->put_Path( _bstr_t(wszCmdLine) );
   if ( FAILED(hr) )   throw com_exception( "Cannot put the action executable path", hr );

   // All done with actionColl, action and execAction - allow to auto-release
}


// Register the task with the task scheduler
void ScheduleTask::_RegisterTask( TaskFolderObj& folder,
                                  TaskDefObj&    task,
                                  const wchar_t* wszTaskName )
{
   if (  wszTaskName == nullptr )   throw runtime_error( "Invalid task name parameter in ScheduleTask::__RegisterTask()" );
   if ( *wszTaskName == L'\0'   )   throw runtime_error( "Blank task name line parameter in ScheduleTask::_RegisterTask()" );

   //RegisterTaskObj  regTask( "IRegisterTask" );        // Don't need this here
   IRegisteredTask* pRegisteredTask  =  NULL;

   HRESULT hr  =  folder.GetObj()->RegisterTaskDefinition( _bstr_t(wszTaskName),
                                                           task.GetObj(),
                                                           TASK_CREATE_OR_UPDATE,
                                                           _variant_t(),
                                                           _variant_t(),
                                                           TASK_LOGON_INTERACTIVE_TOKEN,
                                                           _variant_t( L"" ),
                                                           &pRegisteredTask );

   if ( FAILED(hr) )   throw com_exception( "Error saving the Task", hr );
   pRegisteredTask->Release();
}
