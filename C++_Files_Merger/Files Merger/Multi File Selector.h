#include <Windows.h>
#include <shlwapi.h> // Path related function 
#include <commctrl.h> // common controls
#include "resource.h" 

#pragma comment( lib, "comctl32.lib" )
#pragma comment( lib, "shlwapi.lib" )

void GetFilesPath(HWND hParentWnd,HWND hListBox); 
UINT_PTR CALLBACK _OFNHookProc(HWND hdlg,UINT uiMsg,WPARAM wParam, LPARAM lParam);
LRESULT APIENTRY _OFNSubProc(HWND hwnd,UINT uMsg, WPARAM wParam, LPARAM lParam);
void Add2List(LPSTR lpszFolderPath,LPSTR lpszFiles,HWND hListBox);

/*
This software created by Muhammad Arshad Latti
From Sargodha ( Pakistan) 
e-mail   :: arshadlatti@gmail.com 
web_site :: http://www.geocities.com/arshadlati
*/