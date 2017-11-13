//=============================================================================
// File: App.cpp - Main application entry point
//
//    Copyright (c) 2017 Jeff Reeder
//    All Rights Reserved
//
// This is a raw Win32 "dialog-based" UI app.  It doesn't use MFC, ATL or 
// anything fancy.
//=============================================================================

#include "stdafx.h"

#include "resource.h"
#include "com_exception.h"

#include "Dialog.h"                    // Dialog class


using namespace std;
using namespace PhishMe;


//--------------------------------------------------------------------------
// Application entry-point
//
// NOTE: This is the only symbol in our code that exists in the global scope
//--------------------------------------------------------------------------

int WINAPI _tWinMain( HINSTANCE hInst, 
                      HINSTANCE hPrevInstance, 
                      LPTSTR    lpCmdLine, 
                      int       nCmdShow )
{
   bool bComInitialized  =  false;
   HWND hDlg             =  nullptr;

   try
   {
      // Initialize COM as multi-threaded apartment
      HRESULT hr  =  CoInitializeEx( nullptr, COINIT_MULTITHREADED );

      if ( FAILED(hr) )   throw com_exception( FN_MSG( "Can't initialize COM" ),  hr );

      // Set general COM security levels.
      hr  =  CoInitializeSecurity( nullptr,                          // Default security descriptor
                                   -1,                               // Choose authentication services
                                   nullptr,                          // Choose security providers
                                   nullptr,                          // Reserved
                                   RPC_C_AUTHN_LEVEL_PKT_PRIVACY,    // Authentication type
                                   RPC_C_IMP_LEVEL_IMPERSONATE,      // Impersonate level
                                   nullptr,                          // No authorization list
                                   0,                                // No additional capabilities
                                   nullptr );                        // Reserved

      if ( FAILED(hr) )   throw com_exception( FN_MSG( "Can't initialize COM Security" ),  hr );

      bComInitialized = true;

      //--( Application Code )----------------------------------------------------
      Dialog* pDlg  =  Dialog::GetInstance();

      DialogBox( hInst, 
                 MAKEINTRESOURCE(IDD_DIALOG1), 
                 NULL, 
                 & pDlg->DialogProc );
   }
   catch ( com_exception& ex )
   {
      MessageBox( hDlg,                        // Use ANSI version deliberately
                  Ansi2T( ex.what() ).c_str(),
                  _T("COM Error"),
                  MB_OK | MB_ICONERROR );
   }
   catch ( exception& ex )
   {
      MessageBox( hDlg,                        // Use ANSI version deliberately
                  Ansi2T( ex.what() ).c_str(),
                  _T("Application Error"),
                  MB_OK | MB_ICONERROR );
   }

   if ( bComInitialized )
   {
      // Shutdown COM
      CoUninitialize();
   }

   return 0;
}
