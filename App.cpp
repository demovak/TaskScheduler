//-------------------------------------------
// This is a raw Win32 "dialog-based" UI app.
// It doesn't use MFC, ATL or anything fancy.
//-------------------------------------------

#include "stdafx.h"

#include "resource.h"
#include "Dialog.h"                    // Dialog class
#include "com_exception.h"

// Import the COM objects we'll need
#pragma comment( lib, "taskschd.lib" )
#pragma comment( lib, "comsupp.lib"  )
#pragma comment( lib, "credui.lib"   )


using namespace std;
using namespace PhishMe;


enum ResultCode
{
   resultNormal = 0,
   resultGenericError,
   resultComError
};


//--------------------------------------------------------------------------
// Application entry-point
//
// NOTE: This is the only symbol in our code that exists in the global scope
//--------------------------------------------------------------------------

int __stdcall _tWinMain( HINSTANCE hInst, 
                         HINSTANCE h0, 
                         LPTSTR    lpCmdLine, 
                         int       nCmdShow )
{
   bool bComInitialized  =  false;
   HWND hDlg             =  NULL;
   int  nReturnCode      =  resultNormal;

   try
   {
      // Initialize COM as multi-threaded apartment
      HRESULT hr  =  CoInitializeEx( NULL, COINIT_MULTITHREADED );

      if ( FAILED(hr) )   throw com_exception( "Can't initialize COM" );

      // Set general COM security levels.
      hr  =  CoInitializeSecurity( NULL,                             // Default security descriptor
                                   -1,                               // Choose authentication services
                                   NULL,                             // Choose security providers
                                   NULL,                             // Reserved
                                   RPC_C_AUTHN_LEVEL_PKT_PRIVACY,    // Authentication type
                                   RPC_C_IMP_LEVEL_IMPERSONATE,      // Impersonate level
                                   NULL,                             // No authorization list
                                   0,                                // No additional capabilities
                                   NULL );                           // Reserved

      if ( FAILED(hr) )   throw com_exception( "Can't initialize COM Security", hr );

      //--( Application Code )----------------------------------------------------
      Dialog* pDlg = Dialog::GetInstance();

      // Fire up our Dialog, while using our Dialog class for all the heavy lifting
      HWND hDlg = CreateDialogParam( hInst,
                                     MAKEINTRESOURCE( IDD_DIALOG1 ),
                                     0,
                                     &pDlg->DialogProc,
                                     0 );
      ShowWindow( hDlg, nCmdShow );

      // Handle the boilerplate application message pump
      BOOL nMsgResult;
      MSG  msg;

      while ( ( nMsgResult  =  GetMessage( &msg, 0, 0, 0 ) )  !=  0 )
      {
         if ( nMsgResult == -1 )   throw exception( "GetMessage() failed" );

         if ( !IsDialogMessage( hDlg, &msg ) )
         {
            TranslateMessage( &msg );           // Translate virtual-key messages
            DispatchMessage(  &msg );            // Send it to dialog procedure
         }
      }
   }
   catch ( com_exception& ex )
   {
      nReturnCode = resultComError;
      MessageBoxA( hDlg,                        // Use ANSI version deliberately
                   ex.what(),
                   "COM Error",
                   MB_OK | MB_ICONERROR );
   }
   catch ( exception& ex )
   {
      nReturnCode = resultGenericError;
      MessageBoxA( hDlg,                        // Use ANSI version deliberately
                   ex.what(),
                   "Application Error",
                   MB_OK | MB_ICONERROR );
   }

   if ( bComInitialized )
   {
      // Shutdown COM
      CoUninitialize();
   }

   return nReturnCode;
}
