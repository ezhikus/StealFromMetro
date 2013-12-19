#include "stdafx.h"
#include <iostream>
#include <sstream>
#include <mshtml.h>
#include <atlbase.h>
#include <oleacc.h>
#include "conio.h"

using namespace std;

BOOL CALLBACK EnumChildProc(HWND hwnd,LPARAM lParam)
{
	TCHAR	buf[100];

	::GetClassName( hwnd, (LPTSTR)&buf, 100 );
	if ( _tcscmp( buf, _T("Internet Explorer_Server") ) == 0 )
	{
		*(HWND*)lParam = hwnd;
		return FALSE;
	}
	else
		return TRUE;
};

void GetDocInterface(HWND hWnd) 
{
	CoInitialize( NULL );

	// Explicitly load MSAA so we know if it's installed
	HINSTANCE hInst = ::LoadLibrary( _T("OLEACC.DLL") );
	if ( hInst != NULL )
	{
		if ( hWnd != NULL )
		{
			HWND hWndChild=NULL;
			// Get 1st document window
			::EnumChildWindows( hWnd, EnumChildProc, (LPARAM)&hWndChild );
			if ( hWndChild )
			{
				CComPtr<IHTMLDocument2> spDoc;
				LRESULT lRes;

				UINT nMsg = ::RegisterWindowMessage( _T("WM_HTML_GETOBJECT") );
				::SendMessageTimeout( hWndChild, nMsg, 0L, 0L, SMTO_ABORTIFHUNG, 1000, (DWORD*)&lRes );

				LPFNOBJECTFROMLRESULT pfObjectFromLresult = (LPFNOBJECTFROMLRESULT)::GetProcAddress( hInst, "ObjectFromLresult" );
				if ( pfObjectFromLresult != NULL )
				{
					HRESULT hr;
					hr = (*pfObjectFromLresult)( lRes, IID_IHTMLDocument2, 0, (void**)&spDoc );
					if ( SUCCEEDED(hr) )
					{
						BSTR bstrContent = NULL;
						IHTMLElement *p = 0;
						spDoc->get_body(&p);

						if (p)
						{
							p->get_innerHTML( &bstrContent );
							std::wstring ws(bstrContent, SysStringLen(bstrContent));
							std::string s(ws.begin(), ws.end());
							cout << s;
							p->Release();
						}
					}
				}
			} // else document not ready
		} // else Internet Explorer is not running
		::FreeLibrary( hInst );
	} // else Active Accessibility is not installed
	CoUninitialize();
}

int _tmain(int argc, _TCHAR* argv[])
{
	wstring windowTitle, windowClass;

	wcout << "Please enter parent window title (you can find it by Spy++):" << endl;
	std::getline(std::wcin, windowTitle);
	wcout << "Please enter parent window class (you can find it by Spy++):" << endl;
	std::getline(std::wcin, windowClass);

	HWND hwnd = FindWindow(windowClass.c_str(), windowTitle.c_str());
	wcout << "HWND is " << hwnd << endl;

	try
	{
		GetDocInterface(hwnd);
	}
	catch (...)
	{
	}
	
	_getch();
	return 0;
}

