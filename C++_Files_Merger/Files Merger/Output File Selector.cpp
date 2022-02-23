#include "Output File Selector.h"

void GetOutputPath(LPSTR lpszSaveFile ,HWND hWndParent)
{
	OPENFILENAME ofn;
	char szFileName[MAX_PATH]; 

	ZeroMemory(&ofn, sizeof(ofn));

	lstrcpy(szFileName,lpszSaveFile);

	ofn.lStructSize = sizeof(ofn);
	ofn.hwndOwner = hWndParent;
	ofn.lpstrFile = szFileName;
	ofn.nMaxFile = MAX_PATH;
	ofn.lpstrFilter = "All Files\0*.*\0\0";
	ofn.nFilterIndex = 1;
	ofn.lpstrFileTitle = NULL;
	ofn.nMaxFileTitle = 0;
	ofn.lpstrDefExt ="bin";
	ofn.lpstrInitialDir = NULL;
	ofn.Flags = OFN_HIDEREADONLY;

	if (GetSaveFileName(&ofn)==TRUE)
		lstrcpy(lpszSaveFile,szFileName);// copy file name
	else
		lpszSaveFile[0] = 0 ;// nothing selected return null.
}
// A project by Muhammad Arshad Latti