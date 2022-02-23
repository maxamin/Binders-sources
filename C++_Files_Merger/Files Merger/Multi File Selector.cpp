#include "Multi File Selector.h" 

HWND g_hWndListBox; // List box window handle 
LONG g_lpwpProcOFN; // Original open dialog procedure
BOOL g_bOk = TRUE ;	// Flag value for checking if IDOK or IDCANCEL " if user press enter key or click on open button value will be TRUE if user click on cancel button or press escape key value will be FALSE "
CHAR g_lpszFilesBuffer[65536];	// 64 KB buffer for file names and path.
CHAR g_lpszCurFolderBuffer[MAX_PATH]; // Buffer for folder path 

// Get multiple files names and path into list box  
void GetFilesPath(HWND hParentWnd,HWND hListBox)
{
	OPENFILENAME ofn;
	CHAR lpszFilesPath[MAX_PATH * 8];// Buffer for file names and path but not used in this program maybe simply CHAR lpszFilesPath[MAX_PATH] work fine.

	ZeroMemory(&ofn, sizeof(ofn));
	lpszFilesPath[0] = '\0';

	ofn.lpstrFile = lpszFilesPath;
	ofn.lStructSize = sizeof(ofn);
	ofn.hwndOwner = hParentWnd;// main dialog box window handle 
	ofn.nMaxFile = sizeof(lpszFilesPath);
	ofn.nFilterIndex = 1;
	ofn.lpstrFileTitle = NULL;
	ofn.nMaxFileTitle = 0;
	ofn.lpstrInitialDir = NULL;
	ofn.lpstrFilter ="All Files (*.*)\0*.*\0\0";// user can select any file type.
	ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST   | OFN_HIDEREADONLY | OFN_ALLOWMULTISELECT |OFN_EXPLORER | OFN_ENABLEHOOK ;
	ofn.lpstrTitle="Open...";
	ofn.lpfnHook = _OFNHookProc;// Take control of open dialog box with hooking .
	
	g_hWndListBox = hListBox;// Set list box window handle

		// show open dialog
		GetOpenFileName(&ofn);

}

// Hook procedure
UINT_PTR CALLBACK _OFNHookProc(HWND hdlg,UINT uiMsg,WPARAM wParam, LPARAM lParam)
{


	if ( uiMsg == WM_INITDIALOG )// Before open file dialog box show up.
	{
		g_bOk = TRUE ;// Set ok flag value to TRUE
		g_lpwpProcOFN = SetWindowLong(GetParent(hdlg),GWL_WNDPROC,(LONG)_OFNSubProc);// Subclass the open file dialog box
	}
	return (FALSE);//  Return always false
}

// Subclass procedure of open file dialog box 
LRESULT APIENTRY _OFNSubProc(HWND hwnd,UINT uMsg, WPARAM wParam, LPARAM lParam) 
{ 

    switch (uMsg)
    {
			case WM_COMMAND:
				// Check for IDCANCEL if someone click on cancel button 
				// or press escape key then OK flag value to false.
			if ( LOWORD( wParam ) == IDCANCEL )
				g_bOk = FALSE ;
			//******* Try to make this work on win98
			if ( LOWORD( wParam ) == IDOK )
			{
				g_bOk = TRUE ;
				SendMessage(hwnd,WM_CLOSE,0,0);
			}
			//*******
			break;
			case WM_DESTROY:// Open file dialog box is closing so act and get file names if there.
				// End up subclass of open file dialog. 
				SetWindowLong( hwnd , GWL_WNDPROC, (LONG) g_lpwpProcOFN ); 
				
				if ( g_bOk )// If OK flag is true then get file names and path.
			{
			LPSTR	lpBuffer;
			int cbReqSize = CommDlg_OpenSave_GetSpec(hwnd,NULL,0);// Get length of file names
			
				// Check if g_lpszFilesBuffer ( file names buffer ) is enough for file names
				// if buffer is not enough then allocate new buffer.
				if ( cbReqSize > sizeof(g_lpszFilesBuffer) )
				{
					cbReqSize += MAX_PATH ;
					lpBuffer = (LPSTR) HeapAlloc(GetProcessHeap(),HEAP_ZERO_MEMORY,cbReqSize); //allocate new buffer.
					if (lpBuffer)// check if buffer is allocated.
					{
						// Get current folder path
						CommDlg_OpenSave_GetFolderPath(hwnd,g_lpszCurFolderBuffer, sizeof(g_lpszCurFolderBuffer));
						// Get selected file name(s)
						CommDlg_OpenSave_GetSpec(hwnd,lpBuffer,cbReqSize);
						
						// Add file(s) path  to list box
						Add2List(g_lpszCurFolderBuffer,lpBuffer,g_hWndListBox);
						HeapFree(GetProcessHeap(),0,lpBuffer);// Free allocate buffer
					}
				}
				else// 64 KB buffer is enough
				{
				// Get current folder path
				CommDlg_OpenSave_GetFolderPath(hwnd,g_lpszCurFolderBuffer, sizeof(g_lpszCurFolderBuffer));
				// Get selected file name(s)
				CommDlg_OpenSave_GetSpec(hwnd,g_lpszFilesBuffer,sizeof(g_lpszFilesBuffer));
				
				// Add file(s) path  to list box
				Add2List(g_lpszCurFolderBuffer,g_lpszFilesBuffer,g_hWndListBox);
				}
			}
				break;
	}

	// Always call open file dialog box procedure
    return CallWindowProc( (WNDPROC) g_lpwpProcOFN, hwnd, uMsg, 
        wParam, lParam); 
} 

// Add file(s) name and path to list box.
void Add2List(LPSTR lpszFolderPath,LPSTR lpszFiles,HWND hListBox)
{

	int nFolderLen = lstrlen ( lpszFolderPath) ;
	int nFilesLen  = lstrlen ( lpszFiles ) ;
	char szFullPath[MAX_PATH * 2];
	char szFile [MAX_PATH];

    int n_loop = 0;
	int i_loop = 0;
	int iStartPos,iLen;
	
	if ( ( nFolderLen < 1 ) || ( nFilesLen < 1 )) // check if current folder or selected files string buffer is empty. 
			return;

	if ( lpszFiles[0] != '\"' )// if selected files buffer string not started with ( " ) then user selected only one file.
	{	
		// Simply combine current folder path and file name string buffer to get full path.
		PathCombine(szFullPath,lpszFolderPath,lpszFiles);
		// Add file path to list box
		SendMessage( hListBox, LB_ADDSTRING, NULL, (LPARAM) szFullPath);
		return;// done
	}
	else// Multiple files are selected by user
	{
		while ( lpszFiles[n_loop] != '\0' )// Loop for whole files buffer string.
		{
			
			if ( lpszFiles[n_loop] == '\"' )// if current character is ( " ) that indicate a file name is going to start.
			{
				iStartPos = n_loop + 1;
				i_loop = n_loop + 1;

				while (lpszFiles[i_loop] != '\"')// loop until end for current file name character ( " )
				{
				i_loop++ ;
				if (lpszFiles[i_loop] == 0)
					break;
				}

				n_loop = i_loop;
				iLen = (i_loop - iStartPos) + 1;
				
				lstrcpyn(szFile,&lpszFiles[iStartPos],iLen);// copy file name to temporary  buffer
				PathCombine(szFullPath,lpszFolderPath,szFile);// combine current folder path and file name
				SendMessage( hListBox, LB_ADDSTRING, NULL, (LPARAM) szFullPath);// add file full path to list box
			}		
		n_loop++;
		}
		
	}
}
/*
This software created by Muhammad Arshad Latti
From Sargodha ( Pakistan) 
e-mail   :: arshadlatti@gmail.com 
web_site :: http://www.geocities.com/arshadlati
*/