#include "func.h"
void ErrMsg(HWND hwnd,LPSTR lpstrError)
{
	MessageBox(hwnd,lpstrError,"File Merger-Error",MB_ICONHAND | MB_OK);
}
// get a file size.
// note :: this function taken from "Copy File Part" Application.  
DWORD GetFileSize(LPSTR szFilePath)
{
	DWORD dwFileSize;
	HANDLE hCf;

	// try to open file.
	hCf = CreateFile(szFilePath,
		GENERIC_READ,FILE_SHARE_READ,NULL,
		OPEN_EXISTING,FILE_ATTRIBUTE_NORMAL,
		NULL);

	// check if file is opened.
	if (hCf==INVALID_HANDLE_VALUE)
		return (NULL);// file not open return null.

	dwFileSize=GetFileSize(hCf,NULL);// try to get file size.
	CloseHandle(hCf);// close file.
	
	// check for invalid file size.
	if (dwFileSize==INVALID_FILE_SIZE && GetLastError()!=NO_ERROR)
		return (NULL);// file size is invalid return null.
	
	return (dwFileSize);// return file size.
}
// get percentage in int(integer) from DWORD values.
// note :: this function taken from "Copy File Part" Application.  
int GetPercent(DWORD dwNewValue,DWORD dwMax)
{
	double dMax,dNewValue,dRet;
	int iRet;

	dMax=dwMax;
	dNewValue=dwNewValue;

	dRet =dNewValue / dMax;
	dRet *=100;

	iRet=dRet; // ???

	return(iRet);
}
// A project by Muhammad Arshad Latti