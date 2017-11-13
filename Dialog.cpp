//=============================================================================
// File: Dialog.cpp - Class implementation for main application dialog window
//
//    Copyright (c) 2017 Jeff Reeder
//    All Rights Reserved
//=============================================================================

#include "stdafx.h"

#include "resource.h"

#include "Dialog.h"
#include "ScheduleTask.h"


using namespace std;
using namespace PhishMe;


// STATIC: The only way to get the singleton instance of this dialog
Dialog* Dialog::GetInstance()
{
   static Dialog* pInstance  =  nullptr;

   if ( pInstance == nullptr )
   {
      // Singleton doesn't exist.  Create it
      pInstance  =  new Dialog();
   }

   return pInstance;
}


// Dialog initialization
INT_PTR Dialog::_OnInitDialog( HWND   hDlg,
                               WPARAM wParam,
                               LPARAM lParam )
{
   _hDialog = hDlg;

   // Set number edit box to "5" by default
   SetWindowText( GetDlgItem( hDlg, IDC_TIME_DELAY ), _T("5") );

   return FALSE;
}


INT_PTR Dialog::_Close()
{
   EndDialog( _hDialog, 0 );
   return FALSE;
}


INT_PTR Dialog::_OnCommand( WPARAM wParam,
                            LPARAM lParam )
{
   switch ( LOWORD(wParam) )
   {
      case IDC_ADD_TASK:

         if ( HIWORD(wParam) == BN_CLICKED )
         {
            _OnAddTask();
         }

         break;

      case IDCANCEL:   return _Close();
   }

   return FALSE;
}


void Dialog::_OnAddTask()
{
   try
   {
      TCHAR szCmdLine[256];

      GetWindowText( GetDlgItem( _hDialog, IDC_CMD_LINE ), 
                     szCmdLine,
                     NUM_ELEMENTS(szCmdLine) );

      if ( _tcslen(szCmdLine) == 0 )   throw invalid_argument( FN_MSG( "Can't add an empty task" ) );

      TCHAR szDelay[256];
      GetWindowText( GetDlgItem( _hDialog, IDC_TIME_DELAY ),
                     szDelay,
                     NUM_ELEMENTS(szDelay) );

      if ( _tcslen(szDelay) == 0 )   throw runtime_error( FN_MSG( "Can't add an task without a delay" ) );

      // Get our delay as a numeric value
      long lSeconds  =  _ttol(szDelay);

      if ( lSeconds <= 0L )   throw range_error( FN_MSG( "Task delay cannot be zero or negative" ) );

      // Try to schedule the task
      _AddScheduledTask( szCmdLine, lSeconds );
   }
   catch ( exception& ex )
   {
      _ErrorMsg( ex.what() );   
   }
}


// Add the task to Task Scheduler
void Dialog::_AddScheduledTask( LPCTSTR szCmdLine,
                                long    lDelay )
{
   try
   {
      ScheduleTask scheduler;

      scheduler.CreateScheduledTask( _T("PhishMe Task Scheduler Sample"),
                                     szCmdLine, 
                                     lDelay );

      MessageBox( _hDialog,
                  _T("The task has been scheduled successfully"),
                  _T("Operation Complete" ),
                  MB_OK );
   }
   catch ( exception& ex )
   {
      // Something seriously went wrong inside ScheduleTask
      _ErrorMsg( ex.what(), "Task Creation Failure" );
   }
}


// Display a generic error message (non-fatal)
void Dialog::_ErrorMsg( LPCSTR szText,
                        LPCSTR szCaption  /* = nullptr */ )
{
   tstring sMsg  =  Ansi2T(szText);
   tstring sCap  =  Ansi2T( ( szCaption != nullptr )
                                  ? szCaption
                                  : "Add Task Failure" );

   MessageBox( _hDialog,                             
               Ansi2T(szText).c_str(),
               Ansi2T( ( szCaption != nullptr )
                          ? szCaption
                          : "Add Task Failure" ).c_str(),
               MB_OK | MB_ICONERROR );
}


//=============================================================================
//==[ MESSAGE PUMP SUPPORT ]===================================================
//=============================================================================

// The dialog's real message handler
INT_PTR Dialog::_MessageHandler( HWND   hDlg,
                                 UINT   uMsg,
                                 WPARAM wParam,
                                 LPARAM lParam )
{
   switch ( uMsg )
   {
      case WM_INITDIALOG:   return _OnInitDialog( hDlg, wParam, lParam );
      case WM_COMMAND:      return _OnCommand(          wParam, lParam );
      case WM_CLOSE:        return _Close();

      //case WM_DESTROY:

      //   PostQuitMessage(0);
      //   return TRUE;
   }

   return FALSE;
}


// STATIC: "wedge" to invoke a real class method for the message pump
INT_PTR CALLBACK Dialog::DialogProc( HWND   hDlg,
                                     UINT   uMsg,
                                     WPARAM wParam,
                                     LPARAM lParam )
{
   return GetInstance()->_MessageHandler( hDlg, 
                                          uMsg, 
                                          wParam, 
                                          lParam );
}
