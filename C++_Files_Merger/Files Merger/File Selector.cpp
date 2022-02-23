#include "File Selector.h"

void GetFilePath(HWND hwnd,HWND hListBox)
{
	OPENFILENAME ofn;
	char szFile[MAX_PATH];

	ZeroMemory(&ofn, sizeof(ofn));
	szFile[0]=0;

	ofn.lpstrFile = szFile;
	ofn.lStructSize = sizeof(ofn);
	ofn.hwndOwner = hwnd;
	ofn.lpstrFile[0] = '\0';
	ofn.nMaxFile = sizeof(szFile);
	ofn.nFilterIndex = 1;
	ofn.lpstrFileTitle = NULL;
	ofn.nMaxFileTitle = 0;
	ofn.lpstrInitialDir = NULL;
	ofn.lpstrFilter ="All Files\0*.*\0\0";
	ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST |OFN_HIDEREADONLY;
	ofn.lpstrTitle="Open...";

		if (GetOpenFileName(&ofn)==TRUE)
			SendMessage( hListBox, LB_ADDSTRING, NULL, (LPARAM) szFile);
	
}
// A project by Muhammad Arshad Latti