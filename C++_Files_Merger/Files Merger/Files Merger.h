#include <Windows.h>
//#include <shlwapi.h> // Path related function 
#include <commctrl.h> // common controls
#include "res/resource.h" 

//#pragma comment( lib, "comctl32.lib" )
//#pragma comment( lib, "shlwapi.lib" )
#define MERGE_BUFFER 2048

struct FMPARAM
{
	HWND  hwndMainDlg;
	HWND  hwndListBox;
	HWND  hwndMerging;
	LPSTR lpDestinationFilePath;
	LPSTR lpFileNameBuffer;
	DWORD dwTotalFilesSize;

};

INT_PTR CALLBACK AboutDlgProc(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK DlgProc(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam);
void ListBoxMoveItem(HWND hwndListBox,BOOL bUp);
void StartMerge(HWND hwndDialog,HWND hwndListBox);
LRESULT ListBoxGetMaxStringLen(HWND hwndListBox);
DWORD GetTotalSize(HWND hwndListBox,LPSTR lpBuffer);
DWORD WriteToFile(HANDLE hFile,HANDLE hfIn,DWORD dwTotalFilesSize,DWORD dwPreWritten,DWORD dwCurrentFileSize, HWND hwndMerging,HWND hwndPBFile,HWND hwndPBTotal);
void MergeFiles(FMPARAM * fmParam);
void FillFileNameFromListBox(LPSTR lpBuffer,HWND hwndListBox);
// A project by Muhammad Arshad Latti