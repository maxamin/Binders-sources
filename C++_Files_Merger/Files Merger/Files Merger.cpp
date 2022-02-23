#include "Files Merger.h"
#include "File Selector.h" 
#include "Multi File Selector.h" 
#include "Output File Selector.h"
#include "Messages.h"
#include "func.h"

HMODULE g_hLibShell ;
HINSTANCE g_hInstance;
BOOL g_bStop,g_bPause;
HICON g_hIcon;

// Microsoft® Windows™ application entry point
int APIENTRY  WinMain(HINSTANCE hInstance,HINSTANCE hPrevInstance,LPSTR lpCmdLine,int nCmdShow)
{
	g_hInstance = hInstance;
	// Initializes the common control for Microsoft® Windows™ XP themes(styles) stuff. 
	InitCommonControls();
	g_hLibShell = LoadLibrary("Shell32.dll");
	g_hIcon = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_ICON1) );
	// Show the main dialog box
 	DialogBox(hInstance, (LPCTSTR)IDD_MAIN, 0, (DLGPROC)DlgProc);
	
	DestroyIcon(g_hIcon);
	// exit
	return 0;
}

// Main dialog box messages handler 
LRESULT CALLBACK DlgProc(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{

	switch (uMsg)
	{
case WM_COMMAND:
		if (HIWORD(wParam) == BN_CLICKED)// On button click 
		{
			switch(LOWORD(wParam))
			{
				case IDC_BTN_MERGE:
					StartMerge(hwndDlg,GetDlgItem(hwndDlg,IDC_LISTBOX));
				break;
				case IDC_BTN_ABOUT:// About button is clicked
				DialogBox(g_hInstance,MAKEINTRESOURCE(IDD_ABOUT),hwndDlg,(DLGPROC)AboutDlgProc);
				//MessageBox(hwndDlg,MSG_ABOUT,"About",MB_OK | MB_ICONINFORMATION);
					break;
				case IDC_BTN_ADDFILE:
					GetFilePath(hwndDlg,GetDlgItem(hwndDlg,IDC_LISTBOX));
				break;

				case IDC_BTN_ADDFILES: // Add Files… button is clicked
				// Get multiple file names and path into list box. 
				GetFilesPath(hwndDlg,GetDlgItem(hwndDlg,IDC_LISTBOX));
				break;

				case IDC_BTN_REMOVE: // Remove button is clicked
				if ( SendMessage (GetDlgItem(hwndDlg,IDC_LISTBOX), LB_GETCURSEL, 0, 0) != LB_ERR)
						SendMessage (GetDlgItem(hwndDlg,IDC_LISTBOX), LB_DELETESTRING , SendMessage (GetDlgItem(hwndDlg,IDC_LISTBOX), LB_GETCURSEL, 0, 0), 0);
				break;

				case IDC_BTN_UP:
					 ListBoxMoveItem(GetDlgItem(hwndDlg,IDC_LISTBOX),TRUE);
				break;
				case IDC_BTN_DOWN:
					 ListBoxMoveItem(GetDlgItem(hwndDlg,IDC_LISTBOX),FALSE);
				break;
				case IDC_BTN_CLEAR:
					 SendMessage (GetDlgItem(hwndDlg,IDC_LISTBOX),LB_RESETCONTENT, 0, 0);
				break;
			}
		}
		break;
	case WM_INITDIALOG:
		{	
		SendMessage(hwndDlg,WM_SETICON,ICON_BIG,(LPARAM)g_hIcon);			
			return TRUE;
		}
	case WM_CLOSE:
		EndDialog(hwndDlg,NULL);
			break;
	}
	return FALSE;
}
INT_PTR CALLBACK AboutDlgProc(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	switch (uMsg)
	{
	case WM_COMMAND:
		if ( wParam == IDC_CMD_OK)
			EndDialog(hwndDlg,0);
			break;
	case WM_INITDIALOG:
		SendMessage(hwndDlg,WM_SETICON,ICON_BIG,(LPARAM)g_hIcon);
		return true;
	case WM_CLOSE:
		EndDialog(hwndDlg,0);
		break;
	}
	return FALSE;
}

LRESULT CALLBACK MergingDlgProc(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{

	switch (uMsg)
	{
	case WM_COMMAND:
		switch(LOWORD(wParam))
			{
		case IDC_BTN_STOP:
			g_bStop = TRUE;
			EnableWindow(GetDlgItem(hwndDlg,IDC_BTN_PAUSE),FALSE);
			EnableWindow(GetDlgItem(hwndDlg,IDC_BTN_STOP),FALSE);
			break;
		case IDC_BTN_PAUSE:
				if (g_bPause == FALSE)
				{
				Animate_Stop(GetDlgItem(hwndDlg,IDC_AVI_MERGING));
				g_bPause = TRUE;
				SetDlgItemText(hwndDlg,IDC_BTN_PAUSE,"Resume");
				}
				else
				{
				Animate_Play(GetDlgItem(hwndDlg,IDC_AVI_MERGING),0,-1,-1);
				g_bPause = FALSE;
				SetDlgItemText(hwndDlg,IDC_BTN_PAUSE,"Pause");
				}
			break;
			}
		break;
	case WM_INITDIALOG:			
			// open copying avi from shell32.dll
			Animate_OpenEx(GetDlgItem(hwndDlg,IDC_AVI_MERGING),g_hLibShell,MAKEINTRESOURCE(161));				
			return TRUE;
	case WM_CLOSE:// close button clicked on title bar.
			g_bStop=TRUE;
			break;	
	}
	return FALSE;
}

void ListBoxMoveItem(HWND hwndListBox,BOOL bUp)
{

	LPSTR szSecBuffer = NULL;
	int cbTextLen;
	int iNewIndex;

	if ( SendMessage (hwndListBox, LB_GETCURSEL, 0, 0) != LB_ERR)
	{
			cbTextLen = SendMessage (hwndListBox, LB_GETTEXTLEN, SendMessage (hwndListBox, LB_GETCURSEL, 0, 0), 0);
	
			szSecBuffer = (LPSTR) HeapAlloc(GetProcessHeap(),HEAP_ZERO_MEMORY,cbTextLen + 1);
			if (szSecBuffer)
			{
	        if (SendMessage (hwndListBox, LB_GETTEXT, SendMessage (hwndListBox, LB_GETCURSEL, 0, 0),(LPARAM) szSecBuffer) ==  LB_ERR)
			{
			    HeapFree(GetProcessHeap(),0,szSecBuffer);
				return;
			}
			}
	
		if (bUp == TRUE)
		{
			if (SendMessage (hwndListBox, LB_GETCURSEL, 0, 0) == 0 || SendMessage (hwndListBox, LB_GETCOUNT, 0, 0) < 2)
				 {
			    HeapFree(GetProcessHeap(),0,szSecBuffer);
				return;
			     }
		 iNewIndex = SendMessage (hwndListBox, LB_GETCURSEL, 0, 0) - 1;
		 SendMessage (hwndListBox, LB_DELETESTRING , SendMessage (hwndListBox, LB_GETCURSEL, 0, 0), 0);
		 SendMessage (hwndListBox,  LB_INSERTSTRING , iNewIndex, (LPARAM)szSecBuffer);
		 SendMessage (hwndListBox,  LB_SETCURSEL ,iNewIndex, 0);
		}
		else
		{
				if ((SendMessage (hwndListBox, LB_GETCURSEL, 0, 0) == SendMessage (hwndListBox, LB_GETCOUNT, 0, 0) -1) || (SendMessage (hwndListBox, LB_GETCOUNT, 0, 0) < 2))
				{
			    HeapFree(GetProcessHeap(),0,szSecBuffer);
				return;
			    }
		 iNewIndex = SendMessage (hwndListBox, LB_GETCURSEL, 0, 0) + 1;
		 SendMessage (hwndListBox, LB_DELETESTRING , SendMessage (hwndListBox, LB_GETCURSEL, 0, 0), 0);
		 SendMessage (hwndListBox,  LB_INSERTSTRING , iNewIndex, (LPARAM)szSecBuffer);
		 SendMessage (hwndListBox,  LB_SETCURSEL ,iNewIndex, 0);
		}
	}

	HeapFree(GetProcessHeap(),0,szSecBuffer);
}

void StartMerge(HWND hwndDialog,HWND hwndListBox)
{
	//...
	LRESULT cbFiles;

	cbFiles = SendMessage (hwndListBox, LB_GETCOUNT, 0, 0);

	if ( cbFiles < 1)
	{
		ErrMsg(hwndDialog,ERR_MSG_NOFILES);
		return;
	}
	
	//...
	char * lpOutFilePath =  new char[MAX_PATH];
	
	FillFileNameFromListBox(lpOutFilePath,hwndListBox);
	GetOutputPath(lpOutFilePath,hwndDialog);
	
	if (lstrlen(lpOutFilePath) < 1)
		return;
	
	//...
	LRESULT cbMaxFile;
	DWORD  dwTotalSize;
	LPSTR lpBuffer;
	HWND hwndMerging;

	cbMaxFile = ListBoxGetMaxStringLen(hwndListBox);
	cbMaxFile++;
	lpBuffer = new char[cbMaxFile];
	
	dwTotalSize = GetTotalSize(hwndListBox,lpBuffer);
		
	hwndMerging = CreateDialog(GetModuleHandle(NULL),MAKEINTRESOURCE(IDD_MERGING),hwndDialog, (DLGPROC)MergingDlgProc);
	
	//!!!!!!!!!!!!
	FMPARAM * param = new FMPARAM;
	DWORD dwThreadID;
	HANDLE hThread;

	param->dwTotalFilesSize		= dwTotalSize;
	param->hwndListBox			= hwndListBox;
	param->hwndMainDlg			= hwndDialog;
	param->hwndMerging			= hwndMerging;
	param->lpDestinationFilePath = lpOutFilePath;
	param->lpFileNameBuffer		= lpBuffer;
	
	g_bStop	 = FALSE;
	g_bPause = FALSE;

	hThread = CreateThread(NULL, 0,
			(LPTHREAD_START_ROUTINE) MergeFiles,
			param, 0, &dwThreadID);

	//!!!!!!!!!!!!

}

LRESULT ListBoxGetMaxStringLen(HWND hwndListBox)
{
	LRESULT cbFiles;
	LRESULT cbMaxLen = 0;
	LRESULT cbTextLen;
	

	cbFiles = SendMessage (hwndListBox, LB_GETCOUNT, 0, 0);

	for (int i = 0;i < cbFiles;i++)
	{
	cbTextLen = SendMessage (hwndListBox, LB_GETTEXTLEN, i, 0);
	if ( cbTextLen > cbMaxLen)
			cbMaxLen = cbTextLen;
	}
	return(cbMaxLen);
}
DWORD GetTotalSize(HWND hwndListBox,LPSTR lpBuffer)
{
	LRESULT cbFiles;
	DWORD   dwTotal = 0;

	cbFiles = SendMessage (hwndListBox, LB_GETCOUNT, 0, 0);

	for (int i = 0;i < cbFiles;i++)
	{
		lpBuffer[0] = 0;
		SendMessage (hwndListBox, LB_GETTEXT,i,(LPARAM) lpBuffer);
		dwTotal += GetFileSize(lpBuffer);
	}
	return(dwTotal);
}
void MergeFiles(FMPARAM * fmParam)
{
	HWND hwndStatic = GetDlgItem(fmParam->hwndMerging,IDC_LBL_FILE);
	HWND hwndPBFile = GetDlgItem(fmParam->hwndMerging,IDC_PB_FILE);
	HWND hwndPBTotal = GetDlgItem(fmParam->hwndMerging,IDC_PB_TOTAL);
	DWORD PreWritten = 0;
	DWORD dwTotalSize = fmParam->dwTotalFilesSize;
	HWND hwnd = fmParam->hwndMerging;

	HANDLE hFile = CreateFile(fmParam->lpDestinationFilePath,
		GENERIC_WRITE,FILE_SHARE_READ,NULL,
		CREATE_NEW,FILE_ATTRIBUTE_NORMAL,
		NULL);
	if (hFile==INVALID_HANDLE_VALUE)
	{
		ErrMsg(fmParam->hwndMerging,ERR_MSG_CREATE);
	}
	else
	{
	//...
	LRESULT cbFiles;
	DWORD   dwFileSize;
	HANDLE	hfIn;
	char lpMsg[MAX_PATH + 80];

	
	cbFiles = SendMessage (fmParam->hwndListBox, LB_GETCOUNT, 0, 0);

	for (int i = 0;i < cbFiles;i++)
	{
		if (g_bStop == TRUE)
			 break;
		fmParam->lpFileNameBuffer[0] = 0;
		SendMessage (fmParam->hwndListBox, LB_GETTEXT,i,(LPARAM)fmParam->lpFileNameBuffer);
		dwFileSize = GetFileSize(fmParam->lpFileNameBuffer);
	
		//'''''''''
	SetWindowText(hwndStatic,fmParam->lpFileNameBuffer);

	hfIn = CreateFile(fmParam->lpFileNameBuffer,
		GENERIC_READ,FILE_SHARE_READ,NULL,
		OPEN_EXISTING,FILE_ATTRIBUTE_NORMAL,
		NULL);
	
	if (hfIn==INVALID_HANDLE_VALUE)// check if file opened.
	{
		lpMsg[0] = 0;
		wsprintf(lpMsg,"Error while opening Source(input) file.\n \" %s \"\n Do you want to continue. ",fmParam->lpFileNameBuffer);
		if (MessageBox(fmParam->hwndMerging,lpMsg,"File Merger-Error",MB_ICONQUESTION | MB_YESNO) == IDNO)
		{
		CloseHandle(hFile);
		break;
		}
		else
		{
			continue;
		}

	}
	//'''''''''
	PreWritten = WriteToFile(hFile,hfIn,dwTotalSize,PreWritten,dwFileSize,hwnd,hwndPBFile,hwndPBTotal);
	CloseHandle(hfIn);
	}

	//...
	}


	CloseHandle(hFile);
	EndDialog(fmParam->hwndMerging,NULL);
	delete [] fmParam->lpFileNameBuffer;
	delete [] fmParam->lpDestinationFilePath;
	delete fmParam;
}
DWORD WriteToFile(HANDLE hFile,HANDLE hfIn,DWORD dwTotalFilesSize,DWORD dwPreWritten,DWORD dwCurrentFileSize, HWND hwndMerging,HWND hwndPBFile,HWND hwndPBTotal)
{
		BOOL bReadSuccess,bWriteSuccess;
		DWORD dwRead,dwWrite,dwTotalWritten,dwCurWritten;
		CHAR pBuffer[MERGE_BUFFER];

	dwWrite = 0;
	dwCurWritten = 0;
	dwTotalWritten= dwPreWritten;

	do
	{
		if (g_bPause == TRUE)
		{
			while (g_bPause == TRUE && g_bStop == FALSE )
					Sleep(1);
		}
	// read file
	bReadSuccess = ReadFile(hfIn,pBuffer,MERGE_BUFFER,&dwRead,NULL);

	if (dwRead > 0)// if we get something from file
	{
	// write file
	bWriteSuccess = WriteFile(hFile,pBuffer,dwRead,&dwWrite,0);
	if (bWriteSuccess == 0)
	{
		if (MessageBox(hwndMerging,"Error while writing to file.\n do you want to continue?","File Merger-File Write Error",MB_ICONQUESTION | MB_YESNO) == IDNO)
		{
			g_bStop = TRUE;
		}

	}
	dwTotalWritten += dwWrite;
	dwCurWritten += dwWrite;
	}

	// update progress bar's
	SendMessage(hwndPBFile,PBM_SETPOS,GetPercent(dwCurWritten,dwCurrentFileSize),0);
	SendMessage(hwndPBTotal,PBM_SETPOS,GetPercent(dwTotalWritten,dwTotalFilesSize),0);
	
	} while (dwRead > 0  && bReadSuccess && g_bStop == FALSE );

	return (dwTotalWritten);

}
void FillFileNameFromListBox(LPSTR lpBuffer,HWND hwndListBox)
{
	LRESULT cbFiles;
	char lpListItem[MAX_PATH];

	lpBuffer[0] = 0;

	cbFiles = SendMessage (hwndListBox, LB_GETCOUNT, 0, 0);
	cbFiles--;
	SendMessage (hwndListBox, LB_GETTEXT,cbFiles,(LPARAM)lpListItem);
  wsprintf(lpBuffer,"%s",PathFindFileName(lpListItem));

}
/*
This software created by Muhammad Arshad Latti
From Sargodha ( Pakistan) 
e-mail   :: arshadlatti@gmail.com 
web_site :: http://www.geocities.com/arshadlati
*/