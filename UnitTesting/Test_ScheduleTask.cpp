//=============================================================================
// File: Test_ScheduleTask.cpp - Unit tests for the ScheduleTask class
//
//    Copyright (c) 2017 Jeff Reeder
//    All Rights Reserved
//=============================================================================

#include "stdafx.h"
#include "CppUnitTest.h"

#include "../com_exception.h"
#include "../ScheduleTask.h"                         // Class to test

using namespace std;
using namespace Microsoft::VisualStudio::CppUnitTestFramework;
using namespace PhishMe;


namespace UnitTesting
{
   TEST_CLASS(Test_ScheduleTask)
   {
      TEST_METHOD(ScheduleTask_CreateTaskService)
      {
         try
         {
            ScheduleTask scheduler;
            auto         service  =  scheduler._CreateTaskService();

            Assert::IsTrue(    service.IsValid() );
            Assert::IsNotNull( service.GetObj()  );
         }
         catch ( exception& ex )
         {
            Assert::Fail( EX_FN_MSG( ex, "Exception thrown" ) );
         }
      }

      TEST_METHOD(ScheduleTask_Connect)
      {
         try
         {
            ScheduleTask scheduler;

            // Throw it an invalid service
            Assert::ExpectException<invalid_argument>( [&] () 
               { 
                  scheduler._Connect( ComWrapper<ITaskService>() );
               } );

            scheduler._Connect( scheduler._CreateTaskService() );

            // Try connecting a second time
            scheduler._Connect( scheduler._CreateTaskService() );
         }
         catch ( exception& ex )
         {
            Assert::Fail( EX_FN_MSG( ex, "Exception thrown" ) );
         }
      }

      TEST_METHOD(ScheduleTask_GetRootFolder)
      {
         try
         {
            ScheduleTask scheduler;

            // Throw it an invalid service
            Assert::ExpectException<invalid_argument>( [&]()
            {
               scheduler._GetRootFolder( ComWrapper<ITaskService>() );
            } );

            auto service  =  scheduler._CreateTaskService();

            // Test getting the root folder without connecting first
            Assert::ExpectException<exception>( [&] () 
               {
                  scheduler._GetRootFolder(service);
               } );

            // Go through the full test with a connection
            scheduler._Connect(service);
            auto rootFolder  =  scheduler._GetRootFolder(service);

            Assert::IsTrue(    rootFolder.IsValid() );
            Assert::IsNotNull( rootFolder.GetObj()  );
         }
         catch ( exception& ex )
         {
            Assert::Fail( EX_FN_MSG( ex, "Exception thrown" ) );
         }
      }

      TEST_METHOD(ScheduleTask_CreateTaskBuilder)
      {
         try
         {
            ScheduleTask scheduler;

            // Throw it an invalid task
            Assert::ExpectException<exception>( [&]()
            {
               scheduler._CreateTaskBuilder( ComWrapper<ITaskService>() );
            } );

            auto service = scheduler._CreateTaskService();

            // NOTE: _CreateTaskBuilder() doesn't require a connection to the
            //       task scheduler, so no need to connect in the test.

            // Go through the full test with a connection
            auto task  =  scheduler._CreateTaskBuilder(service);

            Assert::IsTrue(    task.IsValid() );
            Assert::IsNotNull( task.GetObj()  );
         }
         catch ( exception& ex )
         {
            Assert::Fail( EX_FN_MSG( ex, "Exception thrown" ) );
         }
      }

      TEST_METHOD(ScheduleTask_SetAuthor)
      {
         try
         {
            LPCWSTR      wszAuthor  =  L"Some Author";
            ScheduleTask scheduler;

            // Throw it an invalid task
            Assert::ExpectException<invalid_argument>( [&]()
            {
               scheduler._SetAuthor( ComWrapper<ITaskDefinition>(), wszAuthor );
            } );

            auto service  =  scheduler._CreateTaskService();
            auto task     =  scheduler._CreateTaskBuilder(service);

            // Test it with invalid author input strings
            Assert::ExpectException<invalid_argument>( [&]()
            {
               scheduler._SetAuthor( task, nullptr );
            } );

            Assert::ExpectException<invalid_argument>( [&]()
            {
               scheduler._SetAuthor( task, L"" );
            } );

            // Test getting the root folder without connecting.  Should work fine
            scheduler._SetAuthor( task, wszAuthor );
         }
         catch ( exception& ex )
         {
            Assert::Fail( EX_FN_MSG( ex, "Exception thrown" ) );
         }
      }

      TEST_METHOD( ScheduleTask_SetPrincipal )
      {
         try
         {
            LPCWSTR      wszAuthor = L"SomePrincipal";
            ScheduleTask scheduler;

            auto service  =  scheduler._CreateTaskService();
            auto task     =  scheduler._CreateTaskBuilder(service);

            // Give it it an invalid task
            Assert::ExpectException<invalid_argument>( [&]()
            {
               scheduler._SetPrincipalInfo( ComWrapper<ITaskDefinition>(), 
                                            wszAuthor,
                                            TASK_LOGON_INTERACTIVE_TOKEN,
                                            TASK_RUNLEVEL_LUA );
            } );

            // Give it an invalid author string
            Assert::ExpectException<invalid_argument>( [&]()
            {
               scheduler._SetPrincipalInfo( task,
                                            nullptr,
                                            TASK_LOGON_INTERACTIVE_TOKEN,
                                            TASK_RUNLEVEL_LUA );
            } );

            Assert::ExpectException<invalid_argument>( [&]()
            {
               scheduler._SetPrincipalInfo( task,
                                            L"",
                                            TASK_LOGON_INTERACTIVE_TOKEN,
                                            TASK_RUNLEVEL_LUA );
            } );

            // Try invalid logon type
            Assert::ExpectException<invalid_argument>( [&]()
            {
               scheduler._SetPrincipalInfo( task,
                                            wszAuthor,
                                            static_cast<TASK_LOGON_TYPE>( TASK_LOGON_NONE - 1 ),
                                            TASK_RUNLEVEL_LUA );
            } );

            Assert::ExpectException<invalid_argument>( [&]()
            {
               scheduler._SetPrincipalInfo( task,
                                            wszAuthor,
                                            static_cast<TASK_LOGON_TYPE>( TASK_LOGON_INTERACTIVE_TOKEN_OR_PASSWORD + 1 ),
                                            TASK_RUNLEVEL_LUA );
            } );

            // Try invalid run level
            Assert::ExpectException<invalid_argument>( [&]()
            {
               scheduler._SetPrincipalInfo( task,
                                            wszAuthor,
                                            TASK_LOGON_INTERACTIVE_TOKEN,
                                            static_cast<TASK_RUNLEVEL_TYPE>( TASK_RUNLEVEL_LUA - 1 ) );
            } );

            Assert::ExpectException<invalid_argument>( [&]()
            {
               scheduler._SetPrincipalInfo( task,
                                            wszAuthor,
                                            TASK_LOGON_INTERACTIVE_TOKEN,
                                            static_cast<TASK_RUNLEVEL_TYPE>( TASK_RUNLEVEL_HIGHEST + 1 ) );
            } );

            // Now see if a blue sky case works
            scheduler._SetPrincipalInfo( task,
                                         wszAuthor,
                                         TASK_LOGON_INTERACTIVE_TOKEN,
                                         TASK_RUNLEVEL_LUA );
         }
         catch ( exception& ex )
         {
            Assert::Fail( EX_FN_MSG( ex, "Exception thrown" ) );
         }
      }

      TEST_METHOD(ScheduleTask_SetStartWhenAvailable)
      {
         try
         {
            ScheduleTask scheduler;

            // Throw it an invalid task
            Assert::ExpectException<invalid_argument>( [&]()
            {
               scheduler._SetStartWhenAvailable( ComWrapper<ITaskDefinition>(), VARIANT_TRUE );
            } );

            auto service  =  scheduler._CreateTaskService();
            auto task     =  scheduler._CreateTaskBuilder( service );

            // Try it with invalid bWhenAvailable values
            Assert::ExpectException<invalid_argument>( [&]()
            {
               // VARIANT_TRUE = -1
               scheduler._SetStartWhenAvailable( task, VARIANT_TRUE - 1 );
            } );

            Assert::ExpectException<invalid_argument>( [&]()
            {
               // VARIANT_FALSE = 0
               scheduler._SetStartWhenAvailable( task, VARIANT_FALSE + 1 );
            } );

            scheduler._SetStartWhenAvailable( task, VARIANT_TRUE );
         }
         catch ( exception& ex )
         {
            Assert::Fail( EX_FN_MSG( ex, "Exception thrown" ) );
         }
      }

      TEST_METHOD(ScheduleTask_SetTriggerDelay)
      {
         try
         {
            LPCWSTR      wszTriggerName  =  L"SomeTrigger";
            long         lDelay  =  5L;
            ScheduleTask scheduler;

            // Throw it an invalid task
            Assert::ExpectException<invalid_argument>( [&]()
            {
               scheduler._SetTriggerDelay( ComWrapper<ITaskDefinition>(), 
                                           wszTriggerName,
                                           lDelay );
            } );

            auto service  =  scheduler._CreateTaskService();
            auto task     =  scheduler._CreateTaskBuilder(service);

            // Throw it an invalid trigger name
            Assert::ExpectException<invalid_argument>( [&]()
            {
               scheduler._SetTriggerDelay( task,
                                           nullptr,
                                           lDelay );
            } );

            Assert::ExpectException<invalid_argument>( [&]()
            {
               scheduler._SetTriggerDelay( task,
                                           L"",
                                           lDelay );
            } );

            // Try it with invalid lDelay values
            Assert::ExpectException<invalid_argument>( [&]()
            {
               scheduler._SetTriggerDelay( task, 
                                           wszTriggerName,
                                           0L );
            } );

            Assert::ExpectException<invalid_argument>( [&]()
            {
               scheduler._SetTriggerDelay( task, 
                                           wszTriggerName,
                                           -1L );
            } );

            scheduler._SetTriggerDelay( task, 
                                        wszTriggerName,
                                        lDelay );
         }
         catch ( exception& ex )
         {
            Assert::Fail( EX_FN_MSG( ex, "Exception thrown" ) );
         }
      }

      TEST_METHOD(ScheduleTask_SetCommandLineAction)
      {
         try
         {
            LPCWSTR      wszCommandLine  =  LR"("C:\Windows\System32\Notepad.exe")";
            ScheduleTask scheduler;

            // Throw it an invalid task
            Assert::ExpectException<invalid_argument>( [&]()
            {
               scheduler._SetCommandLineAction( ComWrapper<ITaskDefinition>(), wszCommandLine );
            } );

            auto service  =  scheduler._CreateTaskService();
            auto task     =  scheduler._CreateTaskBuilder(service);

            // Try it with invalid szCmdLine values
            Assert::ExpectException<invalid_argument>( [&]()
            {
               scheduler._SetCommandLineAction( task, nullptr );
            } );

            Assert::ExpectException<invalid_argument>( [&]()
            {
               scheduler._SetCommandLineAction( task, L"" );
            } );

            scheduler._SetCommandLineAction( task, wszCommandLine );
         }
         catch ( exception& ex )
         {
            Assert::Fail( EX_FN_MSG( ex, "Exception thrown" ) );
         }
      }

      // NOTE: We don't go through such an exhaustive test on this method.  As
      //       it is the last part of CreateScheduleTask(), the act of actually
      //       registering a legitimate task will be handled by that unit test.
      //       This test only validates against invalid input data - not legit.
      TEST_METHOD(ScheduleTask_RegisterTask)
      {
         try
         {
            ScheduleTask scheduler;
            LPCTSTR      szTaskName  =  _T("Some sample task");

            auto service  =  scheduler._CreateTaskService();

            scheduler._Connect(service);

            auto rootFolder  =  scheduler._GetRootFolder(service);
            auto task        =  scheduler._CreateTaskBuilder(service);

            // Throw it an invalid folder
            Assert::ExpectException<invalid_argument>( [&]()
            {
               scheduler._RegisterTask( ComWrapper<ITaskFolder>(),
                                        task,
                                        T2Wide(szTaskName).c_str() );
            } );

            // Throw it an invalid task
            Assert::ExpectException<invalid_argument>( [&]()
            {
               scheduler._RegisterTask( rootFolder,
                                        ComWrapper<ITaskDefinition>(),
                                        T2Wide(szTaskName).c_str() );
            } );

            // Test it with invalid task names
            Assert::ExpectException<invalid_argument>( [&]()
            {
               scheduler._RegisterTask( rootFolder,
                                        task,
                                        nullptr );
            } );

            Assert::ExpectException<invalid_argument>( [&]()
            {
               scheduler._RegisterTask( rootFolder,
                                        task,
                                        _T("") );
            } );
         }
         catch ( exception& ex )
         {
            Assert::Fail( EX_FN_MSG( ex, "Exception thrown" ) );
         }
      }

      TEST_METHOD(ScheduleTask_CreateScheduledTask)
      {
         try
         {
            LPCTSTR            szTaskName           =  _T("Some sample task");
            LPCTSTR            szCmdLine            =  _T( R"("C:\Windows\System32\Notepad.exe")" );
            long               lDelay               =  10L;
            LPCTSTR            szAuthorName         =  _T("Some Author");
            VARIANT_BOOL       bStartWhenAvailable  =  VARIANT_TRUE;
            TASK_LOGON_TYPE    nLogonType           =  TASK_LOGON_INTERACTIVE_TOKEN;
            TASK_RUNLEVEL_TYPE nRunLevel            =  TASK_RUNLEVEL_LUA;
            LPCTSTR            szPrincipal          =  _T("SomePrincipal");
            LPCTSTR            szTriggerName        =  _T("SomeTrigger");

            ScheduleTask scheduler;

            // Test invalid task name
            Assert::ExpectException<invalid_argument>( [&]()
            {
               scheduler.CreateScheduledTask( nullptr,
                                              szCmdLine,
                                              lDelay,
                                              szAuthorName,
                                              bStartWhenAvailable,
                                              nLogonType,
                                              nRunLevel,
                                              szPrincipal,
                                              szTriggerName );
            } );

            Assert::ExpectException<invalid_argument>( [&]()
            {
               scheduler.CreateScheduledTask( _T(""),
                                              szCmdLine,
                                              lDelay,
                                              szAuthorName,
                                              bStartWhenAvailable,
                                              nLogonType,
                                              nRunLevel,
                                              szPrincipal,
                                              szTriggerName );
            } );

            // Test invalid command line
            Assert::ExpectException<invalid_argument>( [&]()
            {
               scheduler.CreateScheduledTask( szTaskName,
                                              nullptr,
                                              lDelay,
                                              szAuthorName,
                                              bStartWhenAvailable,
                                              nLogonType,
                                              nRunLevel,
                                              szPrincipal,
                                              szTriggerName );
            } );

            Assert::ExpectException<invalid_argument>( [&]()
            {
               scheduler.CreateScheduledTask( szTaskName,
                                              _T(""),
                                              lDelay,
                                              szAuthorName,
                                              bStartWhenAvailable,
                                              nLogonType,
                                              nRunLevel,
                                              szPrincipal,
                                              szTriggerName );
            } );

            // Test invalid delay
            Assert::ExpectException<invalid_argument>( [&]()
            {
               scheduler.CreateScheduledTask( szTaskName,
                                              szCmdLine,
                                              -1L,
                                              szAuthorName,
                                              bStartWhenAvailable,
                                              nLogonType,
                                              nRunLevel,
                                              szPrincipal,
                                              szTriggerName );
            } );

            Assert::ExpectException<invalid_argument>( [&]()
            {
               scheduler.CreateScheduledTask( szTaskName,
                                              szCmdLine,
                                              0L,
                                              szAuthorName,
                                              bStartWhenAvailable,
                                              nLogonType,
                                              nRunLevel,
                                              szPrincipal,
                                              szTriggerName );
            } );

            // Test invalid author name
            Assert::ExpectException<invalid_argument>( [&]()
            {
               scheduler.CreateScheduledTask( szTaskName,
                                              szCmdLine,
                                              lDelay,
                                              nullptr,
                                              bStartWhenAvailable,
                                              nLogonType,
                                              nRunLevel,
                                              szPrincipal,
                                              szTriggerName );
            } );

            Assert::ExpectException<invalid_argument>( [&]()
            {
               scheduler.CreateScheduledTask( szTaskName,
                                              szCmdLine,
                                              lDelay,
                                              _T(""),
                                              bStartWhenAvailable,
                                              nLogonType,
                                              nRunLevel,
                                              szPrincipal,
                                              szTriggerName );
            } );

            // Test invalid start when available
            Assert::ExpectException<invalid_argument>( [&]()
            {
               scheduler.CreateScheduledTask( szTaskName,
                                              szCmdLine,
                                              lDelay,
                                              szAuthorName,
                                              VARIANT_TRUE - 1,         // VARIANT_TRUE = -1
                                              nLogonType,
                                              nRunLevel,
                                              szPrincipal,
                                              szTriggerName );
            } );

            Assert::ExpectException<invalid_argument>( [&]()
            {
               scheduler.CreateScheduledTask( szTaskName,
                                              szCmdLine,
                                              lDelay,
                                              szAuthorName,
                                              VARIANT_FALSE + 1,         // VARIANT_TRUE = 0
                                              nLogonType,
                                              nRunLevel,
                                              szPrincipal,
                                              szTriggerName );
            } );

            // Test invalid logon type
            Assert::ExpectException<invalid_argument>( [&]()
            {
               scheduler.CreateScheduledTask( szTaskName,
                                              szCmdLine,
                                              lDelay,
                                              szAuthorName,
                                              bStartWhenAvailable,
                                              static_cast<TASK_LOGON_TYPE>( TASK_LOGON_NONE - 1 ),
                                              nRunLevel,
                                              szPrincipal,
                                              szTriggerName );
            } );

            Assert::ExpectException<invalid_argument>( [&]()
            {
               scheduler.CreateScheduledTask( szTaskName,
                                              szCmdLine,
                                              lDelay,
                                              szAuthorName,
                                              bStartWhenAvailable,
                                              static_cast<TASK_LOGON_TYPE>( TASK_LOGON_INTERACTIVE_TOKEN_OR_PASSWORD + 1 ),
                                              nRunLevel,
                                              szPrincipal,
                                              szTriggerName );
            } );

            // Test invalid run level
            Assert::ExpectException<invalid_argument>( [&]()
            {
               scheduler.CreateScheduledTask( szTaskName,
                                              szCmdLine,
                                              lDelay,
                                              szAuthorName,
                                              bStartWhenAvailable,
                                              nLogonType,
                                              static_cast<TASK_RUNLEVEL_TYPE>( TASK_RUNLEVEL_LUA - 1 ),
                                              szPrincipal,
                                              szTriggerName );
            } );

            Assert::ExpectException<invalid_argument>( [&]()
            {
               scheduler.CreateScheduledTask( szTaskName,
                                              szCmdLine,
                                              lDelay,
                                              szAuthorName,
                                              bStartWhenAvailable,
                                              nLogonType,
                                              static_cast<TASK_RUNLEVEL_TYPE>( TASK_RUNLEVEL_HIGHEST + 1 ),
                                              szPrincipal,
                                              szTriggerName );
            } );

            // Test invalid principal
            Assert::ExpectException<invalid_argument>( [&]()
            {
               scheduler.CreateScheduledTask( szTaskName,
                                              szCmdLine,
                                              lDelay,
                                              szAuthorName,
                                              bStartWhenAvailable,
                                              nLogonType,
                                              nRunLevel,
                                              nullptr,
                                              szTriggerName );
            } );

            Assert::ExpectException<invalid_argument>( [&]()
            {
               scheduler.CreateScheduledTask( szTaskName,
                                              szCmdLine,
                                              lDelay,
                                              szAuthorName,
                                              bStartWhenAvailable,
                                              nLogonType,
                                              nRunLevel,
                                              _T(""),
                                              szTriggerName );
            } );

            // Test invalid trigger name
            Assert::ExpectException<invalid_argument>( [&]()
            {
               scheduler.CreateScheduledTask( szTaskName,
                                              szCmdLine,
                                              lDelay,
                                              szAuthorName,
                                              bStartWhenAvailable,
                                              nLogonType,
                                              nRunLevel,
                                              szPrincipal,
                                              nullptr );
            } );

            Assert::ExpectException<invalid_argument>( [&]()
            {
               scheduler.CreateScheduledTask( szTaskName,
                                              szCmdLine,
                                              lDelay,
                                              szAuthorName,
                                              bStartWhenAvailable,
                                              nLogonType,
                                              nRunLevel,
                                              szPrincipal,
                                              _T("") );
            } );

            // Test the blue sky scenario - this should create a task with a 10
            // second delay.  
            //
            // NOTE: We'll need to delete the task afterward to cleanup after 
            //       this test
            scheduler.CreateScheduledTask( szTaskName,
                                           szCmdLine,
                                           lDelay,
                                           szAuthorName,
                                           bStartWhenAvailable,
                                           nLogonType,
                                           nRunLevel,
                                           szPrincipal,
                                           szTriggerName );

            // And now delete the task we just created.  If there 
            // was an existing task before this test was run, it 
            // will be deleted.  But since we're using a test 
            // task name, that really shouldn't be a problem.
            scheduler.DeleteScheduledTask(szTaskName);
         }
         catch ( exception& ex )
         {
            Assert::Fail( EX_FN_MSG( ex, "Exception thrown" ) );
         }
      }

      TEST_METHOD( ScheduleTask_DeleteScheduledTask )
      {
         try
         {
            LPCTSTR      szTaskName = _T( "Some sample task" );
            ScheduleTask scheduler;

            // Throw it an invalid task
            Assert::ExpectException<invalid_argument>( [&]()
            {
               scheduler.DeleteScheduledTask( nullptr );
            } );

            Assert::ExpectException<invalid_argument>( [&]()
            {
               scheduler.DeleteScheduledTask( _T( "" ) );
            } );

            scheduler.DeleteScheduledTask( szTaskName );
         }
         catch ( exception& ex )
         {
            Assert::Fail( EX_FN_MSG( ex, "Exception thrown" ) );
         }
      }
    };
}