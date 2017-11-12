#pragma once

namespace PhishMe
{
   // Singleton dialog class to centralize all dialog box code
   class Dialog
   {
   public:

      // The only way to get the singleton instance of this dialog
      static Dialog* GetInstance();


      // The dialog's message pump wedge callback function
      static INT_PTR CALLBACK DialogProc( HWND   hDlg,
                                          UINT   uMsg,
                                          WPARAM wParam,
                                          LPARAM lParam );


   private:

      HWND _hDialog;                   // The one and only dialog HWND

      Dialog()  = default;             // Singleton - force using GetInstance() instead of 'new'
      ~Dialog() = default;             // Singleton

      // The dialog's real message handler
      INT_PTR _MessageHandler( HWND   hDlg,
                               UINT   uMsg,
                               WPARAM wParam,
                               LPARAM lParam );

      // Dialog initialization
      INT_PTR _OnInitDialog( HWND   hDlg, 
                             WPARAM wParam, 
                             LPARAM lParam );

      // Window message command handler
      INT_PTR _OnCommand( WPARAM wParam,
                          LPARAM lParam );

      // Handle the "Add Task" button clicked event
      void _OnAddTask();

      // Add the task to Task Scheduler
      void _AddScheduledTask( LPCTSTR szCmdLine,
                              long    lDelay );

      // Close the window
      INT_PTR _Close();

      // Display a generic error message (non-fatal)
      void _ErrorMsg( LPCSTR szText,
                      LPCSTR szCaption  =  NULL );
   };
}

